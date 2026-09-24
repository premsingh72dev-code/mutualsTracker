import io
import json
import logging
import os
import tempfile
import time
import unittest
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch

from app_logging import JsonLogFormatter, RequestLoggingMiddleware, configure_logging, request_id_context


class WeeklyLoggingTests(unittest.TestCase):
    def test_rotation_retention_and_private_files(self):
        root = logging.getLogger()
        original_handlers, original_level = root.handlers[:], root.level
        service_state = {name: (logging.getLogger(name).handlers[:], logging.getLogger(name).propagate)
                         for name in ('uvicorn', 'uvicorn.error', 'uvicorn.access')}
        root.handlers = []
        try:
            with tempfile.TemporaryDirectory() as directory, patch.dict(os.environ, {
                'LOG_DIR': directory, 'LOG_RETENTION_WEEKS': '2', 'LOG_LEVEL': 'INFO',
            }):
                logger = configure_logging(directory)
                file_handler = next(h for h in root.handlers if hasattr(h, 'rolloverAt'))
                console = next(h for h in root.handlers if not hasattr(h, 'rolloverAt'))
                root.removeHandler(console)
                console.close()
                rollover = datetime.fromtimestamp(file_handler.rolloverAt, timezone.utc)
                self.assertEqual((rollover.weekday(), rollover.hour, rollover.minute), (0, 0, 0))
                self.assertEqual(file_handler.backupCount, 2)
                logger.info('before rotation')
                file_handler.rolloverAt = int(time.time()) - 1
                logger.info('after rotation')
                archives = list(Path(directory).glob('app.log.*'))
                self.assertEqual(len(archives), 1)
                self.assertIn('before rotation', archives[0].read_text())
                self.assertIn('after rotation', (Path(directory) / 'app.log').read_text())
                self.assertEqual((Path(directory) / 'app.log').stat().st_mode & 0o777, 0o640)
                for date in ('2000-01-03', '2000-01-10', '2000-01-17'):
                    (Path(directory) / ('app.log.' + date)).touch()
                self.assertEqual(len(file_handler.getFilesToDelete()), 2)
                # Repeated setup must not duplicate handlers or log entries.
                configure_logging(directory)
                self.assertEqual(len(root.handlers), 1)
        finally:
            for handler in root.handlers:
                handler.close()
            root.handlers = original_handlers
            root.setLevel(original_level)
            for name, (handlers, propagate) in service_state.items():
                logging.getLogger(name).handlers = handlers
                logging.getLogger(name).propagate = propagate

    def test_connection_password_redaction(self):
        record = logging.LogRecord('test', logging.ERROR, '', 1,
                                   'Failed mongodb://review_user:synthetic-secret@localhost/db', (), None)
        formatted = JsonLogFormatter().format(record)
        self.assertNotIn('synthetic-secret', formatted)
        self.assertNotIn('review_user', formatted)
        self.assertIn('***:***@localhost', formatted)


class RequestLoggingTests(unittest.IsolatedAsyncioTestCase):
    def setUp(self):
        self.output = io.StringIO()
        self.logger = logging.getLogger('mutuals.http')
        self.original = (self.logger.handlers[:], self.logger.level, self.logger.propagate)
        handler = logging.StreamHandler(self.output)
        handler.setFormatter(JsonLogFormatter())
        self.logger.handlers = [handler]
        self.logger.setLevel(logging.INFO)
        self.logger.propagate = False
        self.scope = {'type': 'http', 'method': 'POST', 'path': '/api/example',
                      'query_string': b'token=do-not-log',
                      'headers': [(b'authorization', b'Bearer do-not-log')]}

    def tearDown(self):
        for handler in self.logger.handlers:
            handler.close()
        self.logger.handlers, level, self.logger.propagate = self.original
        self.logger.setLevel(level)

    async def receive(self):
        return {'type': 'http.request', 'body': b'password=do-not-log', 'more_body': False}

    async def test_streamed_response_logs_status_duration_and_request_id(self):
        messages = []
        async def send(message):
            messages.append(message)
        async def app(scope, receive, send):
            await receive()
            await send({'type': 'http.response.start', 'status': 200, 'headers': []})
            await send({'type': 'http.response.body', 'body': b'private-report', 'more_body': True})
            self.assertEqual(self.output.getvalue(), '')
            await send({'type': 'http.response.body', 'body': b'final', 'more_body': False})
        await RequestLoggingMiddleware(app)(self.scope, self.receive, send)
        entry = json.loads(self.output.getvalue())
        self.assertEqual(entry['status_code'], 200)
        self.assertEqual(entry['level'], 'INFO')
        self.assertEqual(entry['method'], 'POST')
        self.assertGreaterEqual(entry['duration_ms'], 0)
        self.assertEqual(dict(messages[0]['headers'])[b'x-request-id'].decode(), entry['request_id'])
        self.assertNotIn('do-not-log', self.output.getvalue())
        self.assertNotIn('private-report', self.output.getvalue())
        self.assertIsNone(request_id_context.get())

    async def test_error_status_levels(self):
        async def send(message): pass
        for status, level in ((401, 'WARNING'), (500, 'ERROR')):
            self.output.seek(0); self.output.truncate(0)
            async def app(scope, receive, send):
                await send({'type': 'http.response.start', 'status': status, 'headers': []})
                await send({'type': 'http.response.body', 'body': b''})
            await RequestLoggingMiddleware(app)(self.scope, self.receive, send)
            entry = json.loads(self.output.getvalue())
            self.assertEqual((entry['status_code'], entry['level']), (status, level))

    async def test_unhandled_exception_is_logged_and_reraised(self):
        async def app(scope, receive, send):
            raise ValueError('synthetic failure')
        async def send(message): pass
        with self.assertRaisesRegex(ValueError, 'synthetic failure'):
            await RequestLoggingMiddleware(app)(self.scope, self.receive, send)
        entries = [json.loads(line) for line in self.output.getvalue().splitlines()]
        self.assertIn('ValueError: synthetic failure', entries[0]['exception'])
        self.assertEqual(entries[-1]['status_code'], 500)
        self.assertEqual(entries[0]['request_id'], entries[-1]['request_id'])
        self.assertIsNone(request_id_context.get())

    async def test_non_http_scope_is_passed_through(self):
        visited = []
        async def app(scope, receive, send): visited.append(scope['type'])
        await RequestLoggingMiddleware(app)({'type': 'lifespan'}, None, None)
        self.assertEqual(visited, ['lifespan'])
        self.assertEqual(self.output.getvalue(), '')


if __name__ == '__main__':
    unittest.main()
