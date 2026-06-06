#!/bin/bash
# Example: Upload a CSV and publish payments

BASE_URL="http://localhost:8000"

# 1. Login
echo "Logging in..."
TOKEN=$(curl -s -X POST "$BASE_URL/auth/login" \
  -H "Content-Type: application/json" \
  -d '{"username": "admin", "password": "YourPassword@123"}' \
  | python3 -c "import sys, json; print(json.load(sys.stdin)['access_token'])")

echo "Token: ${TOKEN:0:20}..."

# 2. Upload CSV
echo "Uploading CSV..."
curl -s -X POST "$BASE_URL/upload-csv/" \
  -H "Authorization: Bearer $TOKEN" \
  -F "file=@statement.csv" \
  -F "auto_publish=false" \
  | python3 -m json.tool

# 3. Check upload dates
echo "Available upload dates:"
curl -s "$BASE_URL/upload-dates/" | python3 -m json.tool

# 4. Publish payments for today
TODAY=$(date +%Y-%m-%d)
echo "Publishing payments for $TODAY..."
curl -s -X POST "$BASE_URL/payments/publish?upload_date=$TODAY&async_mode=false" \
  -H "Authorization: Bearer $TOKEN" \
  | python3 -m json.tool

# 5. Download Excel report
echo "Downloading Excel report..."
curl -s "$BASE_URL/download/payments/excel?upload_date=$TODAY" \
  -H "Authorization: Bearer $TOKEN" \
  -o "payments_${TODAY}.xlsx"
echo "Saved: payments_${TODAY}.xlsx"
