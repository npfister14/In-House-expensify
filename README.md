In‑House Expensify (Flask)
==========================

Simple web app to upload a receipt image and fill in details (amount, attendees, occasion, payment method, date). Submits to Airtable and stores the image locally.

Setup
-----

1) Create a `.env` from `.env.example` and fill values.

2) Install Python deps:

```
pip install -r requirements.txt
```

3) Run the Flask server:

```
python server/app.py
```

Then open http://localhost:3000

Environment variables
---------------------

See `.env.example` for all variables. Required for Airtable:

- AIRTABLE_API_KEY
- AIRTABLE_BASE_ID
- AIRTABLE_TABLE_NAME (defaults to `Expenses`)

Optional storage config:

- `UPLOAD_DIR` (default `uploads`)
 - `DEFAULT_CURRENCY` (default `Euro`) — used if the form doesn't provide a currency so reports aren't empty.

Server:

- PORT (default 3000)
- ALLOWED_ORIGIN (CORS; default `*`)
- PUBLIC_BASE_URL (optional; e.g., your public tunnel URL used to build attachment links for Airtable)
- DEFAULT_CURRENCY (optional; default `Euro`) — used if the form doesn't provide a currency
- REPORT_INCLUDE_STATUSES (optional; default `Done,In-Progress,Under Review`) — which statuses to include in report totals/rows

Airtable table schema
---------------------

Create a table (e.g., `Expenses`) with the following fields:

- Amount (Number / Currency)
- Attendees (Single line text)
- Occasion (Single line text)
- Payment (Single select or Text)
- Date (Date)
- Receipt (Attachment)
 - Currency (Single select or Text) — the app will always set this (normalized or `DEFAULT_CURRENCY`).

Frontend
--------

Static files are served from `public/`. Submits a `multipart/form-data` POST to `/api/expenses`.

Notes
-----

- The app uses Flask (`server/app.py`) as the primary server. A legacy Node/Express file used to exist but has been removed to avoid confusion.


SSO via Cloudflare Access (optional)
------------------------------------
Use Cloudflare Access as a zero‑code SSO layer in front of this app.

High‑level steps:
- Put the app behind a hostname like `https://expenses.yourdomain.com` managed by Cloudflare.
- Create a Cloudflare Zero Trust Access Application (type Self‑Hosted) for that hostname.
- Connect your IdP (Google Workspace, Microsoft Entra ID, Okta).
- Create an Access policy: allow emails from `@yourdomain.com`. Optionally, a stricter policy for admin paths.
- If the server isn’t publicly reachable, run Cloudflare Tunnel (cloudflared) to expose the local port.

Testing identity:
- Hit `GET /api/whoami` on your protected hostname. You should see `{ email, source: "cloudflare-access" }`.
- The app doesn’t enforce roles internally; Access is the gate. If you want, use `Cf-Access-Authenticated-User-Email` to show the signed‑in user or add extra checks.

Suggested policies:
- Rule 1 (admin): Path includes `/api/send-summary` or `/api/seed-samples` → Include Group "Finance".
- Rule 2 (default): Path `/*` → Include emails from `@yourdomain.com`.


Email notifications (optional)
------------------------------

A minimal SMTP sender can email a summary after an expense is created. Configure these env vars:

- SMTP_HOST (e.g., smtp.sendgrid.net, smtp.gmail.com)
- SMTP_PORT (default 587)
- SMTP_USER (username or SMTP auth user; for Gmail use app password)
- SMTP_PASS (password)
- EMAIL_FROM (From address)
- EMAIL_TO (comma-separated recipients)

If any are missing, email sending is skipped. The request will not fail if email fails.

Gmail setup quick steps
-----------------------
1) Enable 2‑Step Verification on your Google account.
2) Create an App password: Google Account → Security → App passwords → App: Mail, Device: Other → Generate.
3) Put the 16‑character password (no spaces) into SMTP_PASS.
4) Set SMTP_HOST=smtp.gmail.com, SMTP_PORT=587, SMTP_USER=your@gmail.com, EMAIL_FROM=your@gmail.com.

Test your email config
----------------------
Start the server and open:

- GET /api/test-email → returns { ok: true } if a test email was sent to EMAIL_TO.

Making receipt URLs reachable by Airtable
----------------------------------------
Airtable fetches attachments from the URL you provide. If you run locally, the `http://localhost:PORT/uploads/...` URL is not reachable by Airtable. Set a public base URL:

- Start a tunnel (e.g., ngrok, Cloudflare Tunnel) to your local port.
- Put the public URL into `PUBLIC_BASE_URL` (e.g., `https://1234-12-34-56-78.ngrok-free.app`).
- New uploads will store attachment URLs using this base, so Airtable can retrieve the image.

Why are my report emails empty?
-------------------------------
The monthly report aggregates only records with Status=Done for the chosen month (based on the `Date` field), and groups by Currency. Previously, records missing a Currency were skipped. Now they are included under an `Unknown` currency bucket. Also, when creating a new expense, the backend will set a normalized Currency from the form or use `DEFAULT_CURRENCY` if none is provided. If you still see empty reports:

- Ensure your base records for that month have `Status = Done`.
- Confirm the `Date` field is set within the target month (the report filters by `Date`).
- If running locally, verify Airtable writes succeed (check for errors in the server logs).
- Optionally widen the report to include non-Done statuses by setting `REPORT_INCLUDE_STATUSES` (e.g., `Done,In-Progress`).
