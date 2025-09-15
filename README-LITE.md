# In-House Expensify — Lite

A minimal Flask server focused on essentials:
- Upload expense with a receipt (local storage only)
- List monthly expenses
- Update status (Done, In-Progress, Under Review)
- Image hashing + duplicate hints in listings

Everything else (PDF generation, email, Cloudinary, seeding, SSO helpers) is omitted for a lean experience.

## Run

```
# optional: choose a different port
export PORT=3000

# run lite server
. .venv/bin/activate
python server/app_lite.py
```

## Endpoints
- GET /healthz
- POST /api/expenses (multipart form-data: image, amount, [name, attendees, occasion, payment_method, date, category, reimburse_to, currency])
- GET /api/expenses?month=YYYY-MM
- POST /api/expenses/<record_id>/status  { status: "Done"|"In-Progress"|"Under Review" }

## Env
- AIRTABLE_API_KEY, AIRTABLE_BASE_ID, AIRTABLE_TABLE_NAME (or AIRTABLE_TABLE_ID / AIRTABLE_URL)
- DEFAULT_CURRENCY (optional, defaults to Euro)
- PUBLIC_BASE_URL (optional, for public file URLs)
- UPLOAD_DIR (optional, defaults to ./uploads)

## Notes
- Receipts are stored under uploads/ and served at /uploads/<filename>.
- Each created record includes a Hash field in Airtable.
- Listings include duplicate_hint and duplicate_count.