"""Server logs and HTTP response metrics with weekly UTC rotation."""
import json
import logging
import os
import re
import time
import uuid
from contextvars import ContextVar
from datetime import datetime, time as wall_time, timezone
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path

request_id_context = ContextVar("request_id", default=None)


class JsonLogFormatter(logging.Formatter):
    def format(self, record):
        entry = {
            "timestamp": datetime.fromtimestamp(record.created, timezone.utc).isoformat(),
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
        }
        request_id = getattr(record, "request_id", None) or request_id_context.get()
        if request_id:
            entry["request_id"] = request_id
        for field in ("method", "path", "status_code", "duration_ms"):
            if hasattr(record, field):
                entry[field] = getattr(record, field)
        if record.exc_info:
            entry["exception"] = self.formatException(record.exc_info)
        # Connection errors must not leak database passwords into logs.
        return re.sub(
            r"(mongodb(?:\+srv)?://)[^@\s\"]+@", r"\1***:***@",
            json.dumps(entry, ensure_ascii=False),
        )


def configure_logging(workspace_dir):
    log_dir = Path(os.getenv("LOG_DIR", str(Path(workspace_dir) / "logs")))
    log_dir.mkdir(parents=True, exist_ok=True, mode=0o750)
    log_path = log_dir / "app.log"
    level = os.getenv("LOG_LEVEL", "INFO").upper()
    if level not in ("DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"):
        raise ValueError("LOG_LEVEL must be DEBUG, INFO, WARNING, ERROR, or CRITICAL")
    retention = int(os.getenv("LOG_RETENTION_WEEKS", "12"))
    if retention < 1:
        raise ValueError("LOG_RETENTION_WEEKS must be at least 1")

    root = logging.getLogger()
    root.setLevel(level)
    if not any(getattr(h, "_mutuals_logging", False) for h in root.handlers):
        # Restrict access to logs, including files recreated after a rollover.
        class PrivateWeeklyHandler(TimedRotatingFileHandler):
            def _open(self):
                fd = os.open(self.baseFilename, os.O_CREAT | os.O_APPEND | os.O_WRONLY, 0o640)
                os.chmod(self.baseFilename, 0o640)
                return os.fdopen(fd, self.mode, encoding=self.encoding, errors=self.errors)

        weekly = PrivateWeeklyHandler(
            log_path, when="W0", interval=1, backupCount=retention,
            encoding="utf-8", utc=True, atTime=wall_time(0, 0),
        )
        console = logging.StreamHandler()
        for handler in (weekly, console):
            handler._mutuals_logging = True
            handler.setFormatter(JsonLogFormatter())
            root.addHandler(handler)

    # Uvicorn lifecycle/errors share the same file. Middleware records HTTP access
    # without logging query strings, cookies, authorization headers, or bodies.
    for name in ("uvicorn", "uvicorn.error", "uvicorn.access"):
        service_logger = logging.getLogger(name)
        service_logger.handlers.clear()
        service_logger.propagate = True
    return logging.getLogger("mutuals.app")


class RequestLoggingMiddleware:
    """Pure ASGI middleware: measures complete responses, including streamed exports."""
    def __init__(self, app):
        self.app = app
        self.logger = logging.getLogger("mutuals.http")

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            return await self.app(scope, receive, send)

        request_id = uuid.uuid4().hex
        token = request_id_context.set(request_id)
        started = time.perf_counter()
        status_code = 500
        failed = False
        fields = {"request_id": request_id, "method": scope["method"], "path": scope["path"]}

        async def logged_send(message):
            nonlocal status_code
            if message["type"] == "http.response.start":
                status_code = message["status"]
                headers = [(k, v) for k, v in message.get("headers", []) if k.lower() != b"x-request-id"]
                message = {**message, "headers": headers + [(b"x-request-id", request_id.encode("ascii"))]}
            await send(message)

        try:
            await self.app(scope, receive, logged_send)
        except Exception:
            failed = True
            self.logger.exception("Unhandled request error", extra=fields)
            raise
        finally:
            fields.update(status_code=status_code, duration_ms=round((time.perf_counter() - started) * 1000, 2))
            level = logging.ERROR if failed or status_code >= 500 else logging.WARNING if status_code >= 400 else logging.INFO
            self.logger.log(level, "HTTP response failed" if failed else "HTTP response completed", extra=fields)
            request_id_context.reset(token)
