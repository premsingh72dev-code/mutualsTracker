# Mutual Fund & Dynamic Excel Analytics Platform

A high-performance financial analytics web application designed for mutual fund analysis. Ingests Excel and CSV datasets (Risk Ratios, 1Y/3Y/5Y Rolling Returns), computes Sharpe Ratios, standard deviation, Treynor, and Information Ratios, benchmarks outperformance, builds curated fund baskets, and provides exportable Excel reports.

---

## 🚀 Quick Start with Docker & Docker Compose (Recommended)

Docker provides a reproducible, pre-configured environment with Python 3.10 and MongoDB.

### 1. Start the Stack

```bash
mkdir -p logs
chmod 2750 logs
docker compose up -d --build
```

This starts:
- **FastAPI Analytics App** on `http://localhost` (port `80`)
- **MongoDB Database** on the same server, published only at `127.0.0.1:27017` for SSH tunnels
- **Database files** persisted in `./data/mongo` on the server, mounted at `/data/db` in MongoDB
- **Weekly server logs** persisted in `./logs` on the server, mounted at `/app/logs` in the app

The app connects internally to `mongodb://mongo:27017/` and uses the
`mutual_funds_db` database. Accounts, working sessions, saved analyses, and fund
snapshots are all stored in this instance. Container restarts and recreation
retain the database directory.

### 2. Create the First Admin Account

Because public signup is disabled by default (`ALLOW_SIGNUP=0`), create your first account inside the container:

```bash
docker compose exec app python main.py create-user admin@example.com "YourSecurePassword" "Admin Name"
```

### 3. Open the Dashboard

Navigate to [http://localhost](http://localhost) on the server, or `http://YOUR_SERVER_IP`
from your computer, and sign in with the credentials created above.

### 4. Useful Docker Commands

- **View real-time logs:**
  ```bash
  docker compose logs -f app
  ```
- **Stop containers:**
  ```bash
  docker compose down
  ```
- **Start the existing database and app again:**
  ```bash
  docker compose up -d
  ```

`docker compose down` (including `down -v`) does not erase the bind-mounted
`./data/mongo` directory. Keep that directory to retain your data. After a Python
code change, rebuild/restart the app with `docker compose up -d --build app`;
automatic development reload is disabled for the server process.

### 5. Connect MongoDB Compass through SSH

On your computer, add a new Compass connection with this URI:

```text
mongodb://127.0.0.1:27017/mutual_funds_db?directConnection=true
```

Open **Advanced Connection Options → Proxy / SSH Tunnel**, select
**SSH with Identity File**, and enter:

| Field | Value |
| :--- | :--- |
| SSH Hostname | The server IP or hostname you already use for SSH |
| SSH Port | `22` (or your existing SSH port) |
| SSH Username | `ubuntu` (or your existing server login) |
| SSH Identity File | Your server's private SSH key on your computer |
| SSH Passphrase | Your key's passphrase, if it has one |
| MongoDB authentication | None for this local instance; remote access is authenticated by SSH |

Click **Connect**, then open `mutual_funds_db`. Its app collections are `users`,
`sessions`, `saved_records`, and `mutual_funds`; empty collections may not appear
until first use. Compass and the app access the same server data.

Alternatively, open a tunnel in a terminal on your computer:

```bash
ssh -i /path/to/server-key.pem -N -o ExitOnForwardFailure=yes \
  -L 127.0.0.1:27018:127.0.0.1:27017 ubuntu@YOUR_SERVER_IP
```

Leave that terminal open. In Compass, turn off its built-in SSH tunnel and use:

```text
mongodb://127.0.0.1:27018/mutual_funds_db?directConnection=true
```

Remote database access uses SSH port `22`; MongoDB port `27017` is not published
on public interfaces. Local server processes and the app's Docker network can
also access the database.

See the [MongoDB Compass SSH instructions](https://www.mongodb.com/docs/compass/connect/advanced-connection-options/ssh-connection/).

### 6. Data folder and weekly logs

**Do not remove `data/mongo`.** It contains MongoDB's actual database files, not
an expendable cache. `data/backups` contains database backups, including the
backup made before restricting the MongoDB port. Both are excluded from Git and
Docker builds.

The active log is **`logs/app.log`**. It rotates weekly at **Monday 00:00 UTC**
on the first log event after that time; completed periods are named
`logs/app.log.YYYY-MM-DD`. The default retention is **12 archived weekly files**
plus the active log. Rotation and cleanup happen automatically while the app runs.

Each line is JSON containing the UTC timestamp, severity, logger, and message.
HTTP entries include `request_id`, method, path, `status_code`, and `duration_ms`.
Normal API responses include the same ID in the `X-Request-ID` header. Server
startup, database connection errors, and application exceptions are also logged;
exception entries contain stack traces. Query strings, request/response bodies,
cookies, and authorization headers are not recorded. Logs are server files, not
public web downloads.

```bash
# Watch current requests and errors (on the server)
tail -f logs/app.log

# Find errors and failed HTTP responses
grep '"level": "ERROR"' logs/app.log

# List active and archived weekly logs
ls -lh logs/

# Or read the same JSON logs from the container console
docker compose logs -f app

# Run logging verification tests
python3 -m unittest discover -s tests -v
```

`LOG_RETENTION_WEEKS` controls archive retention and `LOG_LEVEL` controls verbosity.
Set them in the app service's Compose environment and recreate it to apply changes.
Keep `LOG_DIR=/app/logs` in Docker (the default) so logs use the persistent mount.

---

## 🐳 Running Standalone Docker Container

If you prefer to run the application container connecting to a custom MongoDB instance:

```bash
# Build the image
docker build -t mutuals-tracker:latest .

# Run container connected to your MongoDB instance
docker run -d \
  --name mutuals_tracker \
  -p 80:80 \
  -e PORT=80 \
  -e HOST=0.0.0.0 \
  -e MONGO_URI="mongodb://your-mongo-host:27017/" \
  mutuals-tracker:latest

# Create an initial account
docker exec -it mutuals_tracker python main.py create-user you@example.com "password" "Your Name"
```

---

## 💻 Running Locally without Docker

### Prerequisites
- Python 3.10 (recommended)
- pip
- MongoDB running locally, for example with `docker compose up -d mongo`

### Setup
```bash
# Create and activate virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Create an admin user
python main.py create-user admin@example.com "password" "Admin"

# Start development server
python main.py
```

The app will start at `http://127.0.0.1:8005/`.

---

## ⚙️ Environment Variables

| Variable | Default | Description |
| :--- | :--- | :--- |
| `PORT` | `8005` | HTTP port the server listens on |
| `HOST` | `0.0.0.0` | Bind host (`0.0.0.0` for network/LAN access, `127.0.0.1` for local only) |
| `MONGO_URI` | `mongodb://127.0.0.1:27017/` | MongoDB connection string (local or MongoDB Atlas `mongodb+srv://...`) |
| `MONGO_DB_NAME` | `mutual_funds_db` | Name of the MongoDB database |
| `APP_SECRET` | *(auto-generated)* | Key used to sign JWT authentication tokens. Set a persistent key in production |
| `ALLOW_SIGNUP` | `0` | Set to `1` to enable self-registration on the sign-in modal |
| `COOKIE_SECURE` | `0` | Set to `1` when serving behind HTTPS to require Secure cookies |
| `AUTH_TOKEN_TTL_DAYS` | `30` | Duration of authentication sessions in days |
| `ASSET_VERSION` | `8.0.1` | Cache-busting query parameter for CSS and JS assets |
| `LOG_DIR` | `./logs` beside `main.py` | Active and archived weekly log directory (`/app/logs` in Docker) |
| `LOG_LEVEL` | `INFO` | Log verbosity: DEBUG, INFO, WARNING, ERROR, or CRITICAL |
| `LOG_RETENTION_WEEKS` | `12` | Maximum completed weekly log archives to retain |

---

## 📁 Repository Structure

```
├── Dockerfile                  # Production container definition (Python 3.10-slim)
├── docker-compose.yml          # App + MongoDB orchestration
├── .dockerignore               # Docker build context exclusions
├── main.py                     # FastAPI backend, Excel parsing, financial calculations, auth & DB
├── app_logging.py              # Weekly JSON logs and HTTP response timing
├── tests/test_app_logging.py   # Rotation, response logging, and error tests
├── data/                       # Persistent MongoDB files and backups (gitignored)
├── logs/                       # Active and archived weekly server logs (gitignored)
├── requirements.txt            # Pinned runtime dependencies
├── templates/
│   └── index.html              # Single-page dashboard UI (Jinja2)
└── static/
    ├── css/
    │   └── styles.css          # Design system & responsive financial tables
    ├── js/
    │   └── app.js              # Client-side controller, Chart.js integrations & live session state
    └── images/
        └── arp-logo.png        # Application logo
```

---

## 📊 Core Features

- **Excel & CSV Smart Ingestion:** Automatic header detection across legacy `.xls` and modern `.xlsx` sheets.
- **Sharpe Ratio & Rolling Returns Analysis:** Real-time computation of Sharpe Ratios, standard deviation, category averages, and 1Y, 3Y, and 5Y rolling performance.
- **Dynamic Multi-Ratio Screener:** Filter funds by Sharpe, Treynor, Information Ratio, Standard Deviation, and AUM with industry-standard presets.
- **Multi-Device Session Persistence:** Live analysis state and user datasets synced exclusively with MongoDB via PyMongo (stored locally in `./data/mongo`).


### Import accuracy checks (September 2026)

Excel return cells with a percentage number format are read in percentage points:
`0.12` displayed as `12%` imports as `12`. Plain numbers and CSV values are not
rescaled by guessing. Sharpe and other dimensionless ratios retain their original
units. This is supported for XLS, XLSX and XLSM workbooks.

Sequential rolling-return uploads recompute the arithmetic average of all available
1Y, 2Y, 3Y and 5Y horizons, including values from previous uploads. Missing horizons
are excluded; zero and negative returns are retained. This average is not a new CAGR
calculation from NAV history.

Information Ratio column matching excludes unrelated headers such as First NAV Date.
Fund matching rejects conflicting size, index/ETF/FOF and numeric series identifiers,
keeping those funds as separate rows. Existing plan compatibility checks still apply.

Re-upload source spreadsheets for saved analyses created before these fixes; stored
snapshots cannot safely recover original Excel formatting automatically.

Run the regression suite without reading or writing MongoDB data:

```bash
docker cp tests mutuals_tracker_app:/app/tests
docker exec mutuals_tracker_app python -m unittest discover -s tests -v
```

The synthetic XLS fixture is in `tests/fixtures/percentage_returns.xls`; it contains
no investor data. The suite covers import units, missing values, ranks, averages,
benchmarks, category totals, sequential uploads, scheme matching and weekly logs.
