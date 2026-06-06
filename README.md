# FNBB Transaction Processor

> Production FastAPI service that parses FNBB CSV bank statements, classifies transactions as payments vs disbursements, publishes payments to a downstream PHP lending system, and generates Excel/PDF reports.

---

## What This Is

A real production backend built to automate bank reconciliation for a regulated multi-branch lending platform. It eliminates manual CSV processing, ensures payments are correctly classified and published to the lending system, and gives branch managers instant Excel/PDF reports with a single API call.

It is not a tutorial project. It processes real financial data.

---

## Production Deployment

This service runs on an **AWS EC2 instance** with **Nginx as a reverse proxy**. A domain points to the EC2 public IP with SSL terminating at Nginx. The downstream PHP loan management system — hosted on shared hosting — consumes the EC2 APIs over HTTPS.

```
PHP Loan Management System
     (shared hosting)
           │
           │  HTTPS requests to api.yourdomain.com
           ▼
  ┌─────────────────────┐
  │   Nginx (EC2)        │  ← reverse proxy + SSL termination
  │   Port 80/443        │
  └────────┬────────────┘
           │  proxy_pass to localhost:8000
           ▼
  ┌─────────────────────┐
  │  FastAPI / Uvicorn   │  ← application server
  │  (EC2, Port 8000)    │
  └────────┬────────────┘
           │
           ▼
  ┌─────────────────────┐
  │  MySQL + Redis       │  ← data layer (same EC2)
  └─────────────────────┘
```

**Nginx config pattern:**
```nginx
server {
    listen 443 ssl;
    server_name api.yourdomain.com;

    ssl_certificate     /etc/letsencrypt/live/api.yourdomain.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/api.yourdomain.com/privkey.pem;

    location / {
        proxy_pass         http://127.0.0.1:8000;
        proxy_set_header   Host $host;
        proxy_set_header   X-Real-IP $remote_addr;
        proxy_set_header   X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header   X-Forwarded-Proto $scheme;
    }
}

server {
    listen 80;
    server_name api.yourdomain.com;
    return 301 https://$host$request_uri;
}
```

---

## Application Architecture

```
FNBB Bank Statement (CSV)
        │
        ▼
┌───────────────────────────────────────────────────────┐
│            FastAPI Application                        │
│                                                       │
│  ┌─────────────────┐   ┌────────────────────────────┐ │
│  │  CSV Parser      │   │  Auth & RBAC               │ │
│  │  ─────────────  │   │  ──────────────────────    │ │
│  │  Skip 4 headers │   │  Admin / Accountant /      │ │
│  │  Parse amounts  │   │  Collection Officer roles  │ │
│  │  SHA hash dedup │   │  JWT Bearer tokens         │ │
│  │  Flag payments  │   └────────────────────────────┘ │
│  │  vs disbursements│                                 │
│  └─────────────────┘   ┌────────────────────────────┐ │
│                         │  Redis Cache               │ │
│  ┌─────────────────┐   │  ──────────────────────    │ │
│  │  Payment        │   │  Daily summary  300s TTL   │ │
│  │  Publisher      │   │  Transactions    60s TTL   │ │
│  │  ─────────────  │   │  Reports       3600s TTL   │ │
│  │  Exponential    │   └────────────────────────────┘ │
│  │  back-off retry │                                  │
│  │  Publish log    │   ┌────────────────────────────┐ │
│  │  Status track   │   │  Plate Extractor           │ │
│  └─────────────────┘   │  ──────────────────────    │ │
│                         │  Regex plate detection     │ │
│  ┌─────────────────┐   │  Excel VLOOKUP equivalent  │ │
│  │  Report Engine  │   │  Batch plate matching      │ │
│  │  ─────────────  │   └────────────────────────────┘ │
│  │  Excel (.xlsx)  │                                  │
│  │  PDF            │                                  │
│  │  Daily summary  │                                  │
│  └─────────────────┘                                  │
└───────────────────────────────────────────────────────┘
        │                       │
        ▼                       ▼
PHP Lending System        MySQL / MariaDB
(payment publishing)      (transactions, users,
                           publish logs, audit)
```

---

## Key Features

### FNBB CSV Parsing
The FNBB bank statement CSV has 4 header lines before transaction data begins. The parser skips these automatically, handles quoted descriptions containing commas, and classifies each transaction:
- **Positive amount** → Payment (credit / money received)
- **Negative amount** → Disbursement (debit / money paid out)

### SHA-Based Deduplication
Every transaction generates a SHA hash from its date, amount, description, and source file. Re-uploading the same CSV never creates duplicates. The system reports exact counts of new vs duplicate transactions.

### Payment Publishing
Classified payment transactions are published to a downstream PHP lending system via REST API:
- Exponential back-off retry: 1s → 2s → 4s per attempt
- Every attempt logged to `payment_publish_logs` with status, error, and gateway response
- Background (async) or synchronous mode
- Webhook endpoint for PHP system to confirm receipt
- Full publishing status API per transaction

### Vehicle Plate Extraction
Transaction references often contain vehicle registration plate numbers (Botswana format: `BXXXABC`). The plate extractor:
- Scans all transaction descriptions with regex
- Matches found plates against an uploaded Excel registration file
- Returns owner details, amounts paid, and match statistics
- Supports chunked reading for large Excel files (100k+ rows)

### Report Generation
- Excel reports: payments, disbursements, combined — with summary totals
- PDF reports: formatted transaction listings
- Daily summary API: net flow per upload date
- All filterable by upload date, transaction date, or date range

### Security
- JWT Bearer token authentication
- Role-based access: Admin, Accountant, Collection Officer
- Rate limiting: 100/hour global, 5/minute on login
- Security headers middleware (X-Frame-Options, CSP, XSS protection)
- Request size validation (250 MB max)
- All secrets loaded from environment — no hardcoded credentials
- Docs endpoint disabled in production

---

## AWS Migration Path

The service already runs on EC2. Migrating fully to managed AWS services requires minimal changes:

| Current | AWS Managed Equivalent |
|---------|----------------------|
| EC2 + Uvicorn | Keep EC2, or move to Lambda + API Gateway via `handler = Mangum(app)` |
| Nginx on EC2 | AWS Application Load Balancer |
| MySQL on EC2 | Amazon RDS MySQL |
| Redis on EC2 | Amazon ElastiCache Redis |
| Background tasks | Amazon SQS + Lambda |
| CSV upload endpoint | S3 upload trigger → Lambda |
| Reports | S3 storage + CloudFront |
| Rate limiting | API Gateway throttling |

The `handler = Mangum(app)` line at the bottom of `main.py` makes this deployable to AWS Lambda with zero code changes when needed.

---

## Database Tables

| Table | Purpose |
|-------|---------|
| `bank_transactions` | All parsed transactions with payment/disbursement classification |
| `payment_publish_logs` | Publishing history — status, attempt count, PHP response, errors |
| `users` | Application users with roles |

---

## Setup

### Option A — EC2 + Nginx (Production, matches live deployment)

#### 1. Launch EC2 instance
```bash
# Amazon Linux 2023 or Ubuntu 22.04, t2.micro or larger
# Open security group ports: 22 (SSH), 80 (HTTP), 443 (HTTPS)
```

#### 2. Install dependencies on EC2
```bash
# Python 3.11
sudo dnf install python3.11 python3.11-pip -y   # Amazon Linux
# or
sudo apt install python3.11 python3.11-pip -y    # Ubuntu

# Nginx
sudo dnf install nginx -y   # or apt install nginx -y

# MySQL
sudo dnf install mysql-server -y && sudo systemctl start mysqld

# Redis
sudo dnf install redis -y && sudo systemctl start redis
```

#### 3. Clone and configure
```bash
git clone https://github.com/mrmotsumi/fnbb-transaction-processor.git
cd fnbb-transaction-processor
cp app/.env.example app/.env
nano app/.env   # fill in all values
```

Generate secrets:
```bash
python3 -c "import secrets; print(secrets.token_hex(32))"
```

#### 4. Install Python dependencies
```bash
pip3.11 install -r requirements.txt
```

#### 5. Run with systemd (keeps running after SSH disconnect)
```bash
sudo nano /etc/systemd/system/fnbb-api.service
```

```ini
[Unit]
Description=FNBB Transaction Processor
After=network.target

[Service]
User=ec2-user
WorkingDirectory=/home/ec2-user/fnbb-transaction-processor/app
ExecStart=/usr/local/bin/uvicorn main:app --host 127.0.0.1 --port 8000
Restart=always
RestartSec=3
Environment=PYTHONPATH=/home/ec2-user/fnbb-transaction-processor/app

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl daemon-reload
sudo systemctl enable fnbb-api
sudo systemctl start fnbb-api
```

#### 6. Configure Nginx
```bash
sudo nano /etc/nginx/conf.d/fnbb-api.conf
```

```nginx
server {
    listen 80;
    server_name api.yourdomain.com;
    return 301 https://$host$request_uri;
}

server {
    listen 443 ssl;
    server_name api.yourdomain.com;

    ssl_certificate     /etc/letsencrypt/live/api.yourdomain.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/api.yourdomain.com/privkey.pem;

    location / {
        proxy_pass         http://127.0.0.1:8000;
        proxy_set_header   Host $host;
        proxy_set_header   X-Real-IP $remote_addr;
        proxy_set_header   X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header   X-Forwarded-Proto $scheme;
        proxy_read_timeout 300s;
        proxy_connect_timeout 75s;
        client_max_body_size 260M;
    }
}
```

```bash
sudo systemctl reload nginx
```

#### 7. SSL with Let's Encrypt (free)
```bash
sudo dnf install certbot python3-certbot-nginx -y
sudo certbot --nginx -d api.yourdomain.com
```

---

### Option B — Local Development

```bash
git clone https://github.com/mrmotsumi/fnbb-transaction-processor.git
cd fnbb-transaction-processor
cp app/.env.example app/.env
pip install -r requirements.txt
cd app
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

---

### Option C — Docker (Local or EC2)

```bash
docker-compose up -d
```

Includes MySQL and Redis. Edit `app/.env` before starting.

---

## API Endpoints

### Authentication
| Method | Path | Description |
|--------|------|-------------|
| POST | `/auth/login` | Login — returns JWT token |
| GET | `/auth/me` | Current user profile |
| PUT | `/auth/change-password` | Change password |

### Transactions
| Method | Path | Description |
|--------|------|-------------|
| POST | `/upload-csv/` | Upload FNBB CSV statement |
| GET | `/transactions/` | List transactions |
| GET | `/upload-dates/` | All unique upload dates |

### Payments
| Method | Path | Description |
|--------|------|-------------|
| POST | `/payments/publish` | Publish payments to PHP system |
| GET | `/payments/publish/status/{id}` | Publishing status for a transaction |
| GET | `/payments/publish/summary` | Publishing summary stats |
| GET | `/payments/by-plate/{plate}` | Payments by vehicle plate |
| POST | `/payments/by-plates/batch` | Batch plate payment lookup |

### Reports & Downloads
| Method | Path | Description |
|--------|------|-------------|
| GET | `/reports/daily-summary/` | Daily payment/disbursement summary |
| GET | `/reports/payments/` | Payment listing |
| GET | `/reports/disbursements/` | Disbursement listing |
| GET | `/download/payments/excel` | Download payments as Excel |
| GET | `/download/disbursements/excel` | Download disbursements as Excel |
| GET | `/download/payments/pdf` | Download payments as PDF |
| GET | `/download/combined/excel` | Download combined report |

### System
| Method | Path | Description |
|--------|------|-------------|
| GET | `/health/` | Health check (DB + Redis) |
| GET | `/system/status` | System info and uptime |

---

## Environment Variables Reference

| Variable | Required | Description |
|----------|----------|-------------|
| `SECRET_KEY` | ✅ | App signing secret (32+ chars) |
| `JWT_SECRET_KEY` | ✅ | JWT signing secret (32+ chars) |
| `DB_PASSWORD` | ✅ | MySQL password |
| `PHP_API_BASE_URL` | ✅ | Downstream PHP system base URL |
| `PHP_API_KEY` | ✅ | PHP system API key |
| `ADMIN_DEFAULT_PASSWORD` | ✅ | Initial admin password (change after setup) |
| `ENVIRONMENT` | ✅ | `development` / `staging` / `production` |
| `REDIS_URL` | ✅ | Redis connection URL |
| `REPORT_ACCOUNT_NAME` | ✅ | Account name for report headers |
| `REPORT_ACCOUNT_NUMBER` | ✅ | Account number for report headers |

---

## Requirements

- Python 3.11+
- MySQL 5.7+ / MariaDB 10.3+
- Redis 6+
- Nginx (production) or any ASGI host
- AWS EC2 or any Linux VPS, or AWS Lambda via Mangum

---

## License

MIT — free to use, adapt, and build on.

---

*Extracted and anonymised from a production financial platform regulated by a national financial services authority. Account details, credentials, and personal data have been removed. Core business logic and architecture are unchanged.*
