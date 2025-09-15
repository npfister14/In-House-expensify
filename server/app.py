import os
import time
from datetime import date
from pathlib import Path
import random
import hashlib
from io import BytesIO

from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
import sys
import base64
import json
import re
import smtplib
from email.message import EmailMessage

# OpenAI client
from openai import OpenAI

load_dotenv()

# Paths
ROOT_DIR = Path(__file__).resolve().parent.parent
PUBLIC_DIR = ROOT_DIR / "public"
UPLOAD_DIR = ROOT_DIR / (os.getenv("UPLOAD_DIR") or "uploads")
SERVER_DIR = Path(__file__).resolve().parent
if str(SERVER_DIR) not in sys.path:
    sys.path.insert(0, str(SERVER_DIR))
try:
    import reporting  # local module server/reporting.py
except Exception:
    reporting = None

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10 MB

# CORS
allowed_origin = os.getenv("ALLOWED_ORIGIN", "*")
if allowed_origin == "*":
    CORS(app)
else:
    CORS(app, resources={r"/*": {"origins": [allowed_origin]}})

# Local storage only
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

# Airtable via pyairtable
AIRTABLE_API_KEY = os.getenv("AIRTABLE_API_KEY")
AIRTABLE_BASE_ID = os.getenv("AIRTABLE_BASE_ID")
AIRTABLE_TABLE_NAME = os.getenv("AIRTABLE_TABLE_NAME", "Expenses")
AIRTABLE_TABLE_ID = os.getenv("AIRTABLE_TABLE_ID")  # Optional: tblXXXXXXXX
AIRTABLE_URL = os.getenv("AIRTABLE_URL")  # Optional: full Airtable UI URL (we'll parse app/tbl)
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL")  # Optional: e.g., https://your-ngrok-id.ngrok.io
DEFAULT_CURRENCY = os.getenv("DEFAULT_CURRENCY") or "Euro"  # Used when no currency provided

# App-wide allowed categories (match frontend select options)
ALLOWED_CATEGORIES = ["Travels", "Meals", "Supplies", "Others"]
ALLOWED_STATUSES = ["Done", "In-Progress", "Under Review"]
REPORT_INCLUDE_STATUSES = [s.strip() for s in (os.getenv("REPORT_INCLUDE_STATUSES") or "Done,In-Progress,Under Review").split(",") if s.strip()]

# Email config (simple SMTP)
SMTP_HOST = os.getenv("SMTP_HOST")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASS = os.getenv("SMTP_PASS")
EMAIL_FROM = os.getenv("EMAIL_FROM")
EMAIL_TO = os.getenv("EMAIL_TO")  # comma-separated list

# Storage provider
STORAGE_PROVIDER = (os.getenv("STORAGE_PROVIDER") or "local").lower()
CLD_CLOUD_NAME = os.getenv("CLOUDINARY_CLOUD_NAME")
CLD_API_KEY = os.getenv("CLOUDINARY_API_KEY")
CLD_API_SECRET = os.getenv("CLOUDINARY_API_SECRET")
CLD_FOLDER = os.getenv("CLOUDINARY_FOLDER") or "in-house-expensify"

def _configure_cloudinary_if_needed():
    if STORAGE_PROVIDER != "cloudinary":
        return False
    # Require all creds to be present
    if not (CLD_CLOUD_NAME and CLD_API_KEY and CLD_API_SECRET):
        return False
    try:
        import cloudinary
        cloudinary.config(
            cloud_name=CLD_CLOUD_NAME,
            api_key=CLD_API_KEY,
            api_secret=CLD_API_SECRET,
            secure=True,
        )
        return True
    except Exception:
        return False


def _parse_airtable_url(url: str):
    """Extract base (app...) and table (tbl...) IDs from an Airtable UI URL.
    Returns (base_id, table_id) or (None, None) if not found.
    """
    if not url:
        return None, None
    base_match = re.search(r"(app[a-zA-Z0-9]+)", url)
    table_match = re.search(r"(tbl[a-zA-Z0-9]+)", url)
    base_id = base_match.group(1) if base_match else None
    table_id = table_match.group(1) if table_match else None
    return base_id, table_id


def get_airtable_table():
    if not AIRTABLE_API_KEY:
        raise RuntimeError("Missing Airtable env var: AIRTABLE_API_KEY")

    from pyairtable import Table

    parsed_base = parsed_tbl = None
    if AIRTABLE_URL:
        parsed_base, parsed_tbl = _parse_airtable_url(AIRTABLE_URL)
        if not parsed_base:
            raise RuntimeError("AIRTABLE_URL provided but base ID (app...) could not be parsed")

    base_id = AIRTABLE_BASE_ID or parsed_base
    if not base_id:
        raise RuntimeError("Missing Airtable base ID. Set AIRTABLE_BASE_ID or AIRTABLE_URL.")

    table_segment = AIRTABLE_TABLE_ID or parsed_tbl or AIRTABLE_TABLE_NAME
    if not table_segment:
        raise RuntimeError("Missing Airtable table. Set AIRTABLE_TABLE_ID, AIRTABLE_TABLE_NAME, or AIRTABLE_URL.")

    return Table(AIRTABLE_API_KEY, base_id, table_segment)


def build_public_url(path_segment: str) -> str:
    base = (PUBLIC_BASE_URL or request.host_url).rstrip('/')
    if not path_segment.startswith('/'):
        path_segment = '/' + path_segment
    return f"{base}{path_segment}"

# ===== Cloudflare Access helpers =====
def _b64url_decode(data: str) -> bytes:
    import base64
    s = data + '=' * (-len(data) % 4)
    return base64.urlsafe_b64decode(s)

def _cf_email_from_request() -> str | None:
    """Extract user email from Cloudflare Access headers or CF_Authorization cookie.
    Trusts the proxy in front (no JWT signature verification performed here).
    """
    # Preferred direct header if provided by Cloudflare Access
    email = request.headers.get("Cf-Access-Authenticated-User-Email")
    if email:
        return email
    # Sometimes identity is forwarded via X-Forwarded-User
    xf_user = request.headers.get("X-Forwarded-User")
    if xf_user and '@' in xf_user:
        return xf_user
    # Fallback: parse CF_Authorization cookie (JWT) and extract 'email' claim without verifying
    token = request.cookies.get("CF_Authorization") or request.headers.get("Cf-Access-Jwt-Assertion")
    if token and token.count('.') >= 2:
        try:
            parts = token.split('.')
            payload_raw = _b64url_decode(parts[1])
            data = json.loads(payload_raw.decode('utf-8'))
            em = data.get('email') or data.get('sub')
            # Cloudflare often sets country and other claims; ensure it's an email-like value
            if isinstance(em, str) and '@' in em:
                return em
        except Exception:
            pass
    return None

@app.route("/version")
def version():
    return {"version": "1.3.0"}

@app.post("/api/expenses")
def create_expense():
    try:
        # Validate file
        if "image" not in request.files or request.files["image"].filename == "":
            return jsonify({"error": "Image is required"}), 400
        image = request.files["image"]
        # Read bytes once so we can both hash and upload/save
        img_bytes = image.read()
        if not img_bytes:
            return jsonify({"error": "Empty image"}), 400
        # Compute SHA-256 of content
        img_hash = hashlib.sha256(img_bytes).hexdigest()
        # Reset stream for subsequent consumers
        image.stream = BytesIO(img_bytes)

        # Validate amount
        amount_raw = request.form.get("amount")
        if not amount_raw:
            return jsonify({"error": "Amount is required"}), 400
        try:
            amount = float(amount_raw)
        except ValueError:
            return jsonify({"error": "Amount must be a number"}), 400

        attendees = request.form.get("attendees", "")
        occasion = request.form.get("occasion", "")
        payment_method = request.form.get("payment_method", "")
        date_str = request.form.get("date") or date.today().isoformat()
        category = (request.form.get("category") or "Uncategorized").strip()
        reimburse_to = (request.form.get("reimburse_to") or "None").strip() or "None"
        # VAT rate: expect one of 8.1, 2.6, 3.8 (as numbers). If missing or invalid, omit from Airtable.
        def _norm_vat(x):
            try:
                v = float(str(x).strip())
                # If user entered fractional 0.081 etc, scale to percentage numbers
                if 0 < v < 1:
                    v = round(v * 100, 3)
                allowed = {8.1, 2.6, 3.8}
                return v if v in allowed else None
            except Exception:
                return None
        vat_rate = _norm_vat(request.form.get("vat_rate"))

        # Currency: allow only USD, CHF, Euro, CAD
        def _norm_currency(x: str | None) -> str | None:
            if not x:
                return None
            v = str(x).strip().upper()
            # Common aliases and symbols
            if v in {"USD", "US$", "$", "$US", "$USD"}:
                return "USD"
            if v in {"CHF", "SFR", "FR.", "FR", "CHF."}:
                return "CHF"
            if v in {"EUR", "EURO", "€"}:
                return "Euro"
            if v in {"CAD", "C$", "CA$"}:
                return "CAD"
            # If already acceptable
            if v in {"USD", "CHF", "EURO", "CAD"}:
                return "Euro" if v == "EURO" else v
            return None
        currency = _norm_currency(request.form.get("currency"))
        # If not provided, fall back to default currency (ensures reports don't miss entries)
        if not currency:
            currency = DEFAULT_CURRENCY

        name_val = (request.form.get("name") or "").strip()
        if not name_val:
            name_val = f"Expense {date_str}"
        # Generate id as a random 7-digit ID
        random_id = str(random.randint(1_000_000, 9_999_999))

        # Upload/store image
        image_url = None
        attachment = None
        if STORAGE_PROVIDER == "cloudinary" and _configure_cloudinary_if_needed():
            # Upload to Cloudinary
            try:
                from cloudinary import uploader as cld_uploader
                upload_res = cld_uploader.upload(
                    image.stream,
                    folder=CLD_FOLDER,
                    resource_type="image",
                )
                image_url = upload_res.get("secure_url") or upload_res.get("url")
                attachment = [{"url": image_url}] if image_url else None
            except Exception as e:
                return jsonify({"error": f"Cloudinary upload failed: {e}"}), 502
        else:
            # Save image locally
            filename = secure_filename(image.filename)
            name_root, ext = os.path.splitext(filename)
            safe_root = name_root.replace(" ", "_") or "receipt"
            final_name = f"{int(time.time()*1000)}_{safe_root}{ext}"

            save_path = UPLOAD_DIR / final_name
            # Write the bytes we already read
            with open(save_path, "wb") as f:
                f.write(img_bytes)

            # Public URL to the uploaded file
            image_url = build_public_url(f"/uploads/{final_name}")
            attachment = [{"url": image_url}]
        # Write to Airtable
        table = get_airtable_table()
        # Capture uploader email from SSO (if present)
        uploader_email = _cf_email_from_request()

        try:
            payload = {
                "Id": random_id,
                "Name": name_val,
                "Amount": amount,
                "Attendees": attendees,
                "Occasion": occasion,
                "Payment": payment_method,
                "Date": date_str,
                "Date added": date.today().isoformat(),
                "Category": category,
                "Reimburse to": reimburse_to,
                "Status": "Under Review",
                "Receipt": attachment,
                "Hash": img_hash,
            }
            if vat_rate is not None:
                payload["VAT Rate"] = vat_rate
            # Always set a currency (normalized or default)
            payload["Currency"] = currency
            # Attach uploader under the fixed Airtable field name
            if uploader_email:
                payload["Uploaded By"] = uploader_email
            record = table.create(payload, typecast=True)
        except Exception as e:
            # Surface Airtable errors clearly to the client
            return jsonify({
                "error": f"Airtable create failed: {e}",
                "hint": "Check AIRTABLE_URL or AIRTABLE_BASE_ID and table name/id match your base.",
            }), 502

        return jsonify({"ok": True, "recordId": record.get("id"), "imageUrl": image_url, "uploadedBy": uploader_email})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.get("/uploads/<path:filename>")
def uploads(filename):
    # Serve uploaded files
    return send_from_directory(UPLOAD_DIR, filename)


@app.get("/")
def index():
    return send_from_directory(PUBLIC_DIR, "index.html")


@app.get("/styles.css")
def styles():
    return send_from_directory(PUBLIC_DIR, "styles.css")


@app.get("/app.js")
def app_js():
    return send_from_directory(PUBLIC_DIR, "app.js")


@app.get("/records")
def records_page():
    return send_from_directory(PUBLIC_DIR, "records.html")


@app.get("/report")
def report_page():
    return send_from_directory(PUBLIC_DIR, "report.html")


@app.get("/healthz")
def healthz():
    return {"status": "ok"}


@app.errorhandler(404)
def handle_404(e):
    # Ensure API routes return JSON instead of HTML when not found
    if request.path.startswith('/api/'):
        return jsonify({"error": "Not found", "path": request.path}), 404
    return e


@app.errorhandler(405)
def handle_405(e):
    # Ensure API routes return JSON instead of HTML when method not allowed
    if request.path.startswith('/api/'):
        return jsonify({"error": "Method not allowed", "path": request.path}), 405
    return e


def main():
    port = int(os.getenv("PORT", "3000"))
    app.run(host="0.0.0.0", port=port, debug=os.getenv("FLASK_DEBUG") == "1")


@app.get("/api/expenses")
def list_expenses():
    try:
        # month query param as YYYY-MM; default to current month
        month = request.args.get("month")
        if not month:
            today = date.today()
            month = f"{today.year:04d}-{today.month:02d}"

        # Airtable formula to match Date added month
        formula = f"DATETIME_FORMAT({{Date added}}, 'YYYY-MM') = '{month}'"

        table = get_airtable_table()
        # Fetch records for that month
        records = table.all(formula=formula)

        def first_url(attachments):
            try:
                if isinstance(attachments, list) and attachments:
                    return attachments[0].get("url")
            except Exception:
                pass
            return None

        items = []
        for r in records:
            f = r.get("fields", {})
            # Map legacy Approved -> Status if Status missing
            legacy_approved = f.get("Approved")
            status_val = f.get("Status")
            if not status_val and legacy_approved is not None:
                status_val = "Done" if legacy_approved else "Under Review"

            items.append({
                "record_id": r.get("id"),
                "id": f.get("id"),
                "name": f.get("Name"),
                "amount": f.get("Amount"),
                "attendees": f.get("Attendees"),
                "occasion": f.get("Occasion"),
                "payment": f.get("Payment"),
                "date": f.get("Date"),
                "date_added": f.get("Date added"),
                "category": f.get("Category"),
                "reimburse_to": f.get("Reimburse to"),
                "currency": f.get("Currency"),
                "vat_rate": f.get("VAT Rate"),
                "status": status_val or "Under Review",
                "receipt_url": first_url(f.get("Receipt")),
                "hash": f.get("Hash"),
                "uploaded_by": f.get("Uploaded By"),
            })

        # Duplicate detection by hash (ignoring empty/None)
        hash_counts: dict[str, int] = {}
        for it in items:
            h = it.get("hash")
            if h:
                hash_counts[h] = hash_counts.get(h, 0) + 1
        for it in items:
            h = it.get("hash")
            cnt = hash_counts.get(h or "", 0)
            it["duplicate_hint"] = True if h and cnt > 1 else False
            it["duplicate_count"] = cnt if h else 0

        # Sort by Date added desc (client expects this month first and latest first within month)
        items.sort(key=lambda x: (x.get("date_added") or ""), reverse=True)
        return jsonify({"month": month, "count": len(items), "items": items})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.post("/api/expenses/<record_id>/status")
def update_status(record_id):
    try:
        data = request.get_json(silent=True) or request.form
        status = (data.get("status") or "").strip()
        if status not in ALLOWED_STATUSES:
            return jsonify({
                "ok": False,
                "error": "Invalid status",
                "allowed": ALLOWED_STATUSES,
            }), 400

        table = get_airtable_table()
        updated = table.update(record_id, {"Status": status}, typecast=True)
        return jsonify({"ok": True, "recordId": record_id, "status": updated.get("fields", {}).get("Status")})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ===== Email (SMTP minimal) =====
def _email_recipients():
    if not EMAIL_TO:
        return []
    return [addr.strip() for addr in EMAIL_TO.split(",") if addr.strip()]


def send_email(subject: str, html_body: str, to_addrs: list[str], attachments: list[dict] | None = None):
    """Send a simple HTML email using SMTP if configured.
    Raises on failure. Caller may choose to ignore failures.
    """
    if not (SMTP_HOST and EMAIL_FROM and to_addrs):
        # Not configured; silently no-op
        return

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = EMAIL_FROM
    msg["To"] = ", ".join(to_addrs)
    msg.set_content("This email requires an HTML-capable client.")
    msg.add_alternative(html_body, subtype="html")

    # Optional attachments: list of {filename, content(bytes|str), maintype, subtype}
    if attachments:
        for att in attachments:
            if att is None:
                continue
            filename = att.get("filename") or "attachment"
            content = att.get("content")
            maintype = att.get("maintype") or "application"
            subtype = att.get("subtype") or "octet-stream"
            if isinstance(content, str):
                content = content.encode("utf-8")
            if content is None:
                continue
            msg.add_attachment(content, maintype=maintype, subtype=subtype, filename=filename)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=12) as server:
        try:
            server.starttls()
        except Exception:
            # Some servers may not support STARTTLS; continue unencrypted if needed
            pass
        if SMTP_USER:
            server.login(SMTP_USER, SMTP_PASS or "")
        server.send_message(msg)


def maybe_send_email_new_expense(
    *,
    name: str,
    amount: float,
    attendees: str,
    occasion: str,
    payment_method: str,
    date_str: str,
    category: str,
    image_url: str,
    airtable_record_id: str | None,
):
    to = _email_recipients()
    if not to:
        return
    subject = f"New expense: {name or 'Untitled'} — €{amount:.2f}"
    body = f"""
    <h2>New Expense Submitted</h2>
    <ul>
      <li><strong>Name:</strong> {name or '(none)'} </li>
      <li><strong>Amount:</strong> €{amount:.2f}</li>
      <li><strong>Date:</strong> {date_str}</li>
      <li><strong>Category:</strong> {category}</li>
      <li><strong>Payment:</strong> {payment_method}</li>
      <li><strong>Attendees:</strong> {attendees or '(none)'}</li>
      <li><strong>Receipt:</strong> <a href="{image_url}">View</a></li>
      {f'<li><strong>Airtable Record ID:</strong> {airtable_record_id}</li>' if airtable_record_id else ''}
    </ul>
    """
    send_email(subject, body, to)


@app.get("/api/test-email")
def test_email():
    """Send a test email to EMAIL_TO to verify SMTP configuration.
    Returns 400 if not configured.
    """
    to = _email_recipients()
    if not SMTP_HOST or not EMAIL_FROM:
        return jsonify({
            "ok": False,
            "error": "SMTP not configured",
            "hint": "Set SMTP_HOST, EMAIL_FROM, and (optionally) SMTP_PORT/SMTP_USER/SMTP_PASS in .env",
        }), 400
    if not to:
        return jsonify({
            "ok": False,
            "error": "No recipients configured",
            "hint": "Set EMAIL_TO to one or more emails (comma-separated) in .env",
        }), 400

    try:
        subject = "In-House Expensify — Test Email"
        body = """
        <p>This is a test email from the In-House Expensify app.</p>
        <p>If you received this, SMTP is configured correctly.</p>
        """
        send_email(subject, body, to)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


def _normalize_status(s: str | None) -> str:
    v = (s or "").strip()
    # unify hyphen/space variants and case
    lower = v.lower().replace("_", "-").replace(" ", "-")
    if lower in {"done"}:
        return "Done"
    if lower in {"in-progress", "inprogress"}:
        return "In-Progress"
    if lower in {"under-review", "underreview", "under-review."}:
        return "Under Review"
    # Legacy boolean Approved mapping
    if v in {"True", "true", "1"}:
        return "Done"
    if v in {"False", "false", "0"}:
        return "Under Review"
    return v or "Under Review"


def _normalize_payment_method(pm: str | None) -> str:
    v = (pm or "").strip().lower()
    if not v:
        return "Other"
    if v in {"company card", "company-card", "companycard"}:
        return "Company card"
    if v in {"personal", "personal reimbursement", "personal-reimbursement", "reimbursement"}:
        return "Personal"
    if v in {"cash"}:
        return "Cash"
    return pm.strip() if pm else "Other"


def _first_url(attachments):
    try:
        if isinstance(attachments, list) and attachments:
            return attachments[0].get("url")
    except Exception:
        pass
    return None


def _round2(x: float | int | None) -> float:
    try:
        return round(float(x or 0), 2)
    except Exception:
        return 0.0


def _vat_amount(gross: float, vat_rate: float | None) -> float:
    try:
        r = float(vat_rate or 0)
        if 0 < r < 1:
            r = r * 100
        return round(float(gross or 0) * (r / 100.0), 2)
    except Exception:
        return 0.0


def _build_monthly_report(period: str, include_statuses: list[str] | None = None) -> dict:
    """Build monthly report for period 'YYYY-MM' according to the spec.
    Uses Date (occurrence) field for filtering. Ignores records missing Amount or Currency.
    Aggregations include rows with statuses in include_statuses (default from env), while also computing Pending.
    """
    # Which statuses are included in totals/rows
    include_set = set(include_statuses or REPORT_INCLUDE_STATUSES or ["Done"])

    # Fetch records matching period by Date; if Date missing, fall back to Date added
    # Airtable formula: IF(Date, DATETIME_FORMAT(Date,'YYYY-MM')='period', DATETIME_FORMAT({Date added},'YYYY-MM')='period')
    formula = (
        f"IF({{Date}}, DATETIME_FORMAT({{Date}}, 'YYYY-MM') = '{period}', "
        f"DATETIME_FORMAT({{Date added}}, 'YYYY-MM') = '{period}')"
    )
    table = get_airtable_table()
    records = table.all(formula=formula)

    currency_buckets: dict[str, dict] = {}
    rows: list[dict] = []

    for r in records:
        f = r.get("fields", {})
        amount = f.get("Amount")
        currency = f.get("Currency")
        if amount in (None, ""):
            continue  # Ignore records missing Amount
        try:
            gross = float(amount)
        except Exception:
            continue

        cur = str(currency).strip() if currency not in (None, "") else "Unknown"

        status = _normalize_status(f.get("Status") if f.get("Status") is not None else f.get("Approved"))
        payment = _normalize_payment_method(f.get("Payment"))
        payer = (f.get("Payer") or f.get("Reimburse to") or "").strip() or "Unknown"
        category = (f.get("Category") or "").strip() or "Others"
        date_occ = f.get("Date") or f.get("Date added") or ""
        vat_rate = f.get("VAT Rate")

        vat_amt = _vat_amount(gross, vat_rate)
        net_amt = round(gross - vat_amt, 2)

        # Ensure bucket
        b = currency_buckets.setdefault(cur, {
            "totals": {"gross": 0.0, "net": 0.0, "vat": 0.0},
            "byCategory": {},
            "byPaymentMethod": {},
            "companyCardCharged": 0.0,
            "reimbursementsByEmployee": {},
            "pending": {"inProgress": {"count": 0, "gross": 0.0}, "underReview": {"count": 0, "gross": 0.0}},
        })

        # Pending counts include in-month records with In-Progress/Under Review regardless of included statuses
        if status == "In-Progress":
            b["pending"]["inProgress"]["count"] += 1
            b["pending"]["inProgress"]["gross"] += gross
        elif status == "Under Review":
            b["pending"]["underReview"]["count"] += 1
            b["pending"]["underReview"]["gross"] += gross

        # Core aggregations for included statuses
        if status not in include_set:
            continue

        # Totals
        b["totals"]["gross"] += gross
        b["totals"]["net"] += net_amt
        b["totals"]["vat"] += vat_amt

        # By category
        b["byCategory"][category] = b["byCategory"].get(category, 0.0) + gross

        # By payment method
        b["byPaymentMethod"][payment] = b["byPaymentMethod"].get(payment, 0.0) + gross

        # Company card charged
        if payment == "Company card":
            b["companyCardCharged"] += gross

        # Reimbursements by employee (Personal only)
        if payment == "Personal":
            b["reimbursementsByEmployee"][payer or "Unknown"] = b["reimbursementsByEmployee"].get(payer or "Unknown", 0.0) + gross

        # Rows for Done
        rows.append({
            "date": date_occ,
            "payer": payer or "Unknown",
            "category": category,
            "paymentMethod": payment,
            "gross": _round2(gross),
            "net": _round2(net_amt),
            "vat": _round2(vat_amt),
            "currency": cur,
            "status": status,
            "receiptUrl": _first_url(f.get("Receipt")),
            "id": r.get("id"),
        })

    # Round numbers and sort keys
    for cur, b in currency_buckets.items():
        b["totals"] = {k: _round2(v) for k, v in b["totals"].items()}
        # Sort category and payer keys
        b["byCategory"] = {k: _round2(b["byCategory"][k]) for k in sorted(b["byCategory"].keys(), key=lambda x: x.lower())}
        b["byPaymentMethod"] = {k: _round2(b["byPaymentMethod"][k]) for k in sorted(b["byPaymentMethod"].keys(), key=lambda x: x.lower())}
        b["companyCardCharged"] = _round2(b["companyCardCharged"])
        b["reimbursementsByEmployee"] = {k: _round2(b["reimbursementsByEmployee"][k]) for k in sorted(b["reimbursementsByEmployee"].keys(), key=lambda x: x.lower())}
        b["pending"]["inProgress"]["gross"] = _round2(b["pending"]["inProgress"]["gross"])
        b["pending"]["underReview"]["gross"] = _round2(b["pending"]["underReview"]["gross"])

    report = {
        "period": period,
        "currencyBuckets": currency_buckets,
        "rows": rows,
    }
    return report


def _report_rows_to_csv(rows: list[dict]) -> str:
    import csv
    from io import StringIO
    buf = StringIO()
    writer = csv.writer(buf)
    writer.writerow(["date", "payer", "category", "paymentMethod", "gross", "net", "vat", "currency", "status", "receiptUrl", "id"])
    for r in rows:
        writer.writerow([
            r.get("date", ""),
            r.get("payer", "Unknown"),
            r.get("category", ""),
            r.get("paymentMethod", ""),
            f"{_round2(r.get('gross')):.2f}",
            f"{_round2(r.get('net')):.2f}",
            f"{_round2(r.get('vat')):.2f}",
            r.get("currency", ""),
            r.get("status", ""),
            r.get("receiptUrl", ""),
            r.get("id", ""),
        ])
    return buf.getvalue()


def _render_report_pdf(report: dict) -> bytes:
    """Render a polished Monthly Expense Report PDF with charts and styled sections."""
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Flowable
        )
        from reportlab.graphics.shapes import Drawing, String, Circle
        from reportlab.graphics.charts.barcharts import HorizontalBarChart
        from reportlab.graphics.charts.piecharts import Pie
    except Exception as e:
        raise RuntimeError(f"PDF generation requires reportlab. Install it: pip install reportlab ({e})")

    from io import BytesIO
    from datetime import datetime

    # Helpers
    def sum_bucket_totals(currency_buckets: dict):
        gross = net = vat = 0.0
        for b in currency_buckets.values():
            t = b.get('totals', {})
            gross += float(t.get('gross', 0) or 0)
            net += float(t.get('net', 0) or 0)
            vat += float(t.get('vat', 0) or 0)
        return _round2(gross), _round2(net), _round2(vat)

    def make_cards_table(cards: list[tuple[str, str]]):
        # Simple 3-up card layout using a table
        data = [[Paragraph(f"<b>{title}</b><br/>{value}", styles['BodyText']) for title, value in cards]]
        t = Table(data, colWidths=[(doc.width)/len(cards)]*len(cards))
        t.setStyle(TableStyle([
            ('BOX', (0,0), (-1,-1), 0.6, colors.HexColor('#d0d0d0')),
            ('INNERGRID', (0,0), (-1,-1), 0.6, colors.HexColor('#e0e0e0')),
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f7f7f7')),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
            ('RIGHTPADDING', (0,0), (-1,-1), 10),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ]))
        return t

    def dict_table(title_left: str, data_dict: dict, col1w=90*mm, col2w=30*mm):
        rows = [[title_left, 'Amount']]
        for k, v in data_dict.items():
            rows.append([str(k), f"{_round2(v):.2f}"])
        t = Table(rows, colWidths=[col1w, col2w])
        t.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#eaeaea')),
            ('ALIGN', (1,1), (-1,-1), 'RIGHT'),
        ]))
        return t

    def bar_chart_from_dict(title: str, data_dict: dict):
        try:
            if not data_dict:
                return None
            labels = list(data_dict.keys())
            values = [float(data_dict[k] or 0) for k in labels]
            h = 140
            w = 340
            d = Drawing(w, h)
            chart = HorizontalBarChart()
            chart.x = 60
            chart.y = 10
            chart.height = h - 30
            chart.width = w - 80
            chart.data = [values]
            chart.categoryAxis.categoryNames = labels
            chart.valueAxis.valueMin = 0
            chart.bars[0].fillColor = colors.HexColor('#5b8cff')
            d.add(chart)
            d.add(String(0, h-10, title, fontSize=10))
            return d
        except Exception:
            return None

    def donut_from_dict(title: str, data_dict: dict):
        try:
            if not data_dict:
                return None
            labels = list(data_dict.keys())
            values = [float(data_dict[k] or 0) for k in labels]
            size = 160
            d = Drawing(size, size)
            p = Pie()
            p.x = 10
            p.y = 10
            p.width = size-20
            p.height = size-20
            p.data = values
            p.labels = [str(l) for l in labels]
            palette = [
                colors.HexColor('#5b8cff'), colors.HexColor('#7b61ff'), colors.HexColor('#22c55e'),
                colors.HexColor('#f59e0b'), colors.HexColor('#ef4444')
            ]
            for i, s in enumerate(p.slices):
                s.fillColor = palette[i % len(palette)]
            d.add(p)
            # fake donut by overlaying circle
            d.add(Circle(size/2, size/2, 26, fillColor=colors.white, strokeColor=colors.white))
            d.add(String(0, size+2, title, fontSize=10))
            return d
        except Exception:
            return None

    # Build document
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4, leftMargin=16*mm, rightMargin=16*mm, topMargin=16*mm, bottomMargin=16*mm
    )
    styles = getSampleStyleSheet()
    h1 = ParagraphStyle('h1', parent=styles['Heading1'], fontSize=18, leading=22)
    h2 = ParagraphStyle('h2', parent=styles['Heading2'], fontSize=14, leading=18)
    h3 = ParagraphStyle('h3', parent=styles['Heading3'], fontSize=12, leading=16)
    normal = styles['BodyText']

    elems: list = []

    # Header
    ts = datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')
    title = f"Monthly Expense Report — {report.get('period','')}"
    header_table = Table([
        [Paragraph(title, h1), Paragraph('<i>Company Logo</i>', ParagraphStyle('logo', parent=normal, alignment=2))],
        [Paragraph(f"Generated {ts}", normal), '']
    ], colWidths=[doc.width*0.75, doc.width*0.25])
    header_table.setStyle(TableStyle([
        ('ALIGN', (1,0), (1,0), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    elems.append(header_table)
    elems.append(Spacer(1, 6))

    # Executive Summary
    currency_buckets = report.get('currencyBuckets', {}) or {}
    gross, net, vat = sum_bucket_totals(currency_buckets)
    elems.append(Paragraph('Executive Summary', h2))
    elems.append(make_cards_table([
        ('Total Gross', f"{gross:.2f}"),
        ('Total Net', f"{net:.2f}"),
        ('Total VAT', f"{vat:.2f}"),
    ]))
    elems.append(Spacer(1, 6))

    # Totals per currency
    if len(currency_buckets) > 0:
        rows = [['Currency', 'Gross', 'Net', 'VAT']]
        for cur, b in currency_buckets.items():
            t = b.get('totals', {})
            rows.append([cur, f"{_round2(t.get('gross')):.2f}", f"{_round2(t.get('net')):.2f}", f"{_round2(t.get('vat')):.2f}"])
        t = Table(rows, colWidths=[30*mm, 30*mm, 30*mm, 30*mm])
        t.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#eaeaea')),
            ('ALIGN', (1,1), (-1,-1), 'RIGHT'),
        ]))
        elems.append(t)
        elems.append(Spacer(1, 8))

    # Prepare rows by currency for precise computations (e.g., reimbursements Done only)
    all_rows = report.get('rows', [])
    rows_by_cur: dict[str, list] = {}
    for r in all_rows:
        rows_by_cur.setdefault(r.get('currency') or 'Unknown', []).append(r)

    # Breakdowns per currency
    elems.append(Paragraph('Breakdowns', h2))
    for cur, _b in currency_buckets.items():
        elems.append(Paragraph(f'Currency: {cur}', h3))
        b = currency_buckets[cur]
        by_cat = b.get('byCategory', {})
        by_pay = b.get('byPaymentMethod', {})

        # Charts
        bar = bar_chart_from_dict('By Category', by_cat)
        donut = donut_from_dict('By Payment Method', by_pay)
        charts_row = []
        if bar:
            charts_row.append(bar)
        if donut:
            charts_row.append(donut)
        if charts_row:
            cw = [doc.width/len(charts_row)]*len(charts_row)
            charts_table = Table([charts_row], colWidths=cw)
            charts_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ]))
            elems.append(charts_table)
            elems.append(Spacer(1, 4))

        # Tables
        elems.append(dict_table('Category', by_cat))
        elems.append(Spacer(1, 4))
        elems.append(dict_table('Payment Method', by_pay))
        elems.append(Spacer(1, 6))

        # Company card overview (Done only)
        cc_total = 0.0
        for r in rows_by_cur.get(cur, []):
            if (r.get('status') or '').lower() == 'done' and (r.get('paymentMethod') or '').lower().startswith('company'):
                cc_total += float(r.get('gross') or 0)
        elems.append(make_cards_table([('Company Card Spent', f"{_round2(cc_total):.2f}")]))

        # Reimbursements by employee (Done + Personal/Cash)
        reimb: dict[str, float] = {}
        for r in rows_by_cur.get(cur, []):
            pm = (r.get('paymentMethod') or '').lower()
            st = (r.get('status') or '').lower()
            if st == 'done' and (pm == 'personal' or pm == 'cash'):
                key = (r.get('payer') or 'Unknown').strip() or 'Unknown'
                reimb[key] = reimb.get(key, 0.0) + float(r.get('gross') or 0)
        if reimb:
            elems.append(Spacer(1, 4))
            elems.append(Paragraph('Reimbursements owed (Done, Personal/Cash)', h3))
            elems.append(dict_table('Employee', reimb))
        elems.append(Spacer(1, 6))

    # Pending overview (aggregated)
    ip_count = ur_count = 0
    ip_gross = ur_gross = 0.0
    for b in currency_buckets.values():
        p = b.get('pending', {})
        ip = p.get('inProgress', { 'count': 0, 'gross': 0 })
        ur = p.get('underReview', { 'count': 0, 'gross': 0 })
        ip_count += int(ip.get('count', 0) or 0)
        ur_count += int(ur.get('count', 0) or 0)
        ip_gross += float(ip.get('gross', 0) or 0)
        ur_gross += float(ur.get('gross', 0) or 0)
    elems.append(Paragraph('Pending Overview', h2))
    elems.append(make_cards_table([
        ('In‑Progress', f"{ip_count} • {ip_gross:.2f}"),
        ('Under Review', f"{ur_count} • {ur_gross:.2f}"),
    ]))
    elems.append(Spacer(1, 8))

    # Detailed rows
    rows = report.get('rows', [])
    elems.append(Paragraph('Detailed Rows', h2))
    if not rows:
        elems.append(Paragraph('No rows for this period with the selected statuses.', normal))
    else:
        header = ['Date', 'Payer', 'Category', 'Payment', 'Gross', 'Net', 'VAT', 'Currency', 'Status']
        data = [header]
        for r in rows:
            data.append([
                r.get('date',''), r.get('payer',''), r.get('category',''), r.get('paymentMethod',''),
                f"{_round2(r.get('gross')):.2f}", f"{_round2(r.get('net')):.2f}", f"{_round2(r.get('vat')):.2f}",
                r.get('currency',''), r.get('status',''),
            ])
        table = Table(data, repeatRows=1)
        style_cmds = [
            ('GRID', (0,0), (-1,-1), 0.4, colors.lightgrey),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#eaeaea')),
            ('ALIGN', (4,1), (6,-1), 'RIGHT'),
        ]
        # Zebra striping and status chip coloring
        for i in range(1, len(data)):
            bg = colors.whitesmoke if i % 2 == 0 else colors.white
            style_cmds.append(('BACKGROUND', (0,i), (-1,i), bg))
            status_text = data[i][-1].lower()
            status_col = colors.HexColor('#22c55e') if 'done' in status_text else (
                colors.HexColor('#f59e0b') if 'progress' in status_text else colors.HexColor('#ef4444')
            )
            style_cmds.append(('BACKGROUND', (-1,i), (-1,i), status_col))
            style_cmds.append(('TEXTCOLOR', (-1,i), (-1,i), colors.white))
        table.setStyle(TableStyle(style_cmds))
        elems.append(table)

    # Footer with page numbers
    def _footer(canvas, _doc):
        canvas.saveState()
        footer_text = 'In‑House Expensify — Confidential • Page %d' % (_doc.page)
        canvas.setFillColor(colors.HexColor('#666666'))
        canvas.setFont('Helvetica', 9)
        canvas.drawString(16*mm, 12*mm, footer_text)
        canvas.restoreState()

    doc.build(elems, onFirstPage=_footer, onLaterPages=_footer)
    return buf.getvalue()


def generate_report_pdf(period: str, *, include_statuses: list[str] | None = None) -> tuple[bytes, dict]:
    """Build the monthly report for a period and render it to PDF using reporting module."""
    report = _build_monthly_report(period, include_statuses=include_statuses)
    if reporting is None:
        raise RuntimeError("reporting module not available")
    pdf_bytes = reporting.render_report_pdf(report)
    return pdf_bytes, report


def _render_error_pdf(message: str) -> bytes:
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.platypus import SimpleDocTemplate, Paragraph
        from reportlab.lib.units import mm
    except Exception:
        # As a last resort, return bytes of the message; browser will likely show blank
        return (message or 'Error').encode('utf-8')
    from io import BytesIO
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=16*mm, rightMargin=16*mm, topMargin=16*mm, bottomMargin=16*mm)
    styles = getSampleStyleSheet()
    elems = [Paragraph('Report Generation Error', styles['Heading2']), Paragraph(message or 'Unknown error', styles['BodyText'])]
    doc.build(elems)
    return buf.getvalue()


@app.get("/api/expense-report")
def get_expense_report():
    try:
        # Accept month=YYYY-MM or year & month separately
        period = request.args.get("month")
        if not period:
            year = (request.args.get("year") or str(date.today().year)).zfill(4)
            month = (request.args.get("monthNum") or request.args.get("m") or str(date.today().month)).zfill(2)
            period = f"{year}-{month}"
        # Optional: include statuses in totals/rows via query (?statuses=Done,In-Progress,Under Review)
        statuses_param = request.args.get("statuses")
        include_statuses = [s.strip() for s in statuses_param.split(",")] if statuses_param else None

        report = _build_monthly_report(period, include_statuses=include_statuses)
        return jsonify(report)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.get("/api/expense-report.check")
def expense_report_check():
    """Lightweight diagnostics to help troubleshoot report generation in production."""
    checks = {}
    # ReportLab availability
    try:
        import reportlab  # type: ignore
        checks['reportlab'] = getattr(reportlab, '__version__', 'present')
    except Exception as e:
        checks['reportlab'] = f"missing: {e}"
    # Airtable configuration and connectivity (shallow)
    try:
        tbl = get_airtable_table()
        # small query to verify access (bounded)
        recs = tbl.all(max_records=1)
        checks['airtable'] = f"ok ({len(recs)} accessible)"
    except Exception as e:
        checks['airtable'] = f"error: {e}"
    return jsonify(checks)


@app.post("/api/send-summary")
def send_summary_email():
    """Send a detailed monthly report email with CSV attachment for the given month (YYYY-MM)."""
    # Determine period
    period = request.args.get("month") or (request.json.get("month") if request.is_json else request.form.get("month"))
    if not period:
        today = date.today()
        period = f"{today.year:04d}-{today.month:02d}"

    # Build report and PDF attachment
    from datetime import datetime
    try:
        statuses_param = request.args.get("statuses") or ((request.json or {}).get("statuses") if request.is_json else request.form.get("statuses"))
        include_statuses = [s.strip() for s in statuses_param.split(",")] if statuses_param else None
        pdf_bytes, report = generate_report_pdf(period, include_statuses=include_statuses)
        ts = datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
        html = f"""
        <p>Please find the attached PDF summary for {report['period']}.<br/>
        Generated at {ts}.</p>
        """
    except Exception as e:
        msg = str(e)
        if "reportlab" in msg.lower():
            return jsonify({"ok": False, "error": f"PDF render failed: {e}"}), 500
        return jsonify({"ok": False, "error": f"Airtable error: {e}"}), 502
    attachments = [{"filename": f"expenses_{report['period']}.pdf", "content": pdf_bytes, "maintype": "application", "subtype": "pdf"}]

    # Determine recipient: current logged-in user via SSO
    if not (SMTP_HOST and EMAIL_FROM):
        return jsonify({"ok": False, "error": "SMTP not configured (SMTP_HOST/EMAIL_FROM)"}), 400
    current_email = _cf_email_from_request()
    if not current_email:
        return jsonify({"ok": False, "error": "No authenticated user email from SSO"}), 401
    to = [current_email]
    rows = report.get("rows", [])
    subject = f"Expense Report {report['period']} — rows: {len(rows)}"

    # Send asynchronously to avoid blocking the request on slow SMTP
    try:
        import threading
        threading.Thread(target=lambda: send_email(subject, html, to, attachments=attachments), daemon=True).start()
    except Exception:
        # If threading fails, fall back to inline (still with SMTP timeout set in send_email)
        try:
            send_email(subject, html, to, attachments=attachments)
        except Exception as e:
            return jsonify({"ok": False, "error": str(e)}), 500

    total = sum((r.get('gross') or 0) for r in rows)
    return jsonify({"ok": True, "month": report['period'], "count": len(rows), "total": _round2(total), "queued": True})


@app.get("/api/expense-report.pdf")
def get_expense_report_pdf():
    """Return the monthly report as a PDF for download/preview."""
    try:
        period = request.args.get("month")
        if not period:
            year = (request.args.get("year") or str(date.today().year)).zfill(4)
            month = (request.args.get("monthNum") or request.args.get("m") or str(date.today().month)).zfill(2)
            period = f"{year}-{month}"
        statuses_param = request.args.get("statuses")
        include_statuses = [s.strip() for s in statuses_param.split(",")] if statuses_param else None
        pdf, _report = generate_report_pdf(period, include_statuses=include_statuses)
        from flask import Response
        resp = Response(pdf, mimetype='application/pdf')
        resp.headers['Content-Disposition'] = f"inline; filename=expenses_{period}.pdf"
        # Prevent stale caching in browsers/proxies
        resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        resp.headers['Pragma'] = 'no-cache'
        resp.headers['Expires'] = '0'
        return resp
    except Exception as e:
        # Return a small error PDF so the viewer doesn't show a blank page
        from flask import Response
        if reporting is not None:
            pdf = reporting.render_error_pdf(f"{e}")
        else:
            pdf = (str(e) or 'Error').encode('utf-8')
        resp = Response(pdf, mimetype='application/pdf', status=500)
        resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        resp.headers['Pragma'] = 'no-cache'
        resp.headers['Expires'] = '0'
        return resp


@app.get("/api/whoami")
def whoami():
    """Return identity info from upstream SSO proxy headers (e.g., Cloudflare Access).
    Useful to verify SSO configuration and show the signed-in user in the UI.
    """
    # Common headers from Cloudflare Access
    email = _cf_email_from_request() or ""
    user = email or request.headers.get("X-Forwarded-User") or ""
    # Optionally include groups or JWT info if forwarded (disabled by default for security)
    source = "cloudflare-access" if email else "unauthenticated"
    return jsonify({
        "email": email or user,
        "source": source,
        "headersPresent": [h for h in [
            "Cf-Access-Authenticated-User-Email",
            "X-Forwarded-User",
        ] if request.headers.get(h)]
    })


# ===== OpenAI Receipt Analysis =====
def _openai_client():
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("Missing OPENAI_API_KEY")
    return OpenAI(api_key=api_key)


@app.post("/api/analyze")
def analyze_receipt():
    try:
        if "image" not in request.files or request.files["image"].filename == "":
            return jsonify({"error": "Image is required"}), 400
        file = request.files["image"]
        img_bytes = file.read()
        mime = file.mimetype or "image/jpeg"
        b64 = base64.b64encode(img_bytes).decode("ascii")
        data_url = f"data:{mime};base64,{b64}"

        # Instruction and JSON schema
        instruction = (
            "Extract fields from the receipt image and return STRICT JSON only. Keys: "
            "amount (number), attendees (string, comma-separated names), occasion (string), "
            "payment_method (string: Company Card | Personal Reimbursement | Cash | Other), "
            "date (YYYY-MM-DD), category (string), name (string), vat_rate (number), currency (string). "
            "For category, choose ONE of exactly: Travels, Meals, Supplies, Others. "
            "For name, suggest a short descriptive title (e.g., '<Merchant> <context>' like 'Starbucks Team Lunch'). "
            "For vat_rate, ALWAYS choose ONE of exactly: 8.1, 2.6, 3.8 per Swiss VAT. "
            "Heuristics: accommodation/hotel = 3.8; grocery/food retail/books/news/water/medicines = 2.6; restaurant/cafe/hospitality services = 8.1; alcohol always 8.1. "
            "If a VAT percentage is printed, use it (mapping 8 or 8.1 -> 8.1; 2.5/2.6 -> 2.6; 3.7/3.8 -> 3.8). If not printed, infer from merchant/store type and products, and if ambiguous default to 8.1. "
            "For currency, choose ONE of exactly: USD, CHF, Euro, CAD; do not use other values. If unknown, use empty string. "
            "If unknown, use empty string for strings and 0 for amount."
        )

        client = _openai_client()
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            temperature=0,
            response_format={"type": "json_object"},
            messages=[
                {
                    "role": "system",
                    "content": "You are a precise receipt extraction engine. Output strict JSON only.",
                },
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": instruction},
                        {"type": "image_url", "image_url": {"url": data_url}},
                    ],
                },
            ],
        )

        content = resp.choices[0].message.content or "{}"
        try:
            data = json.loads(content)
        except Exception:
            data = {}

        # Normalize
        def as_float(x):
            try:
                return float(x)
            except Exception:
                return 0.0

        def normalize_category(value: str) -> str:
            v = (value or "").strip().lower()
            if not v:
                return ""
            travel_syn = [
                "travel",
                "travels",
                "trip",
                "transport",
                "transportation",
                "flight",
                "train",
                "taxi",
                "uber",
                "lyft",
                "car",
            ]
            meals_syn = [
                "meal",
                "meals",
                "food",
                "lunch",
                "dinner",
                "breakfast",
                "restaurant",
                "cafe",
                "coffee",
            ]
            supplies_syn = [
                "supply",
                "supplies",
                "office supplies",
                "stationery",
                "hardware",
                "equipment",
            ]
            if v in travel_syn or any(w in v for w in travel_syn):
                return "Travels"
            if v in meals_syn or any(w in v for w in meals_syn):
                return "Meals"
            if v in supplies_syn or any(w in v for w in supplies_syn):
                return "Supplies"
            # If the model already returned an allowed exact value, keep it
            for allowed in ALLOWED_CATEGORIES:
                if v == allowed.lower():
                    return allowed
            return "Others"

        raw_category = (data.get("category") or "").strip()
        normalized_category = normalize_category(raw_category)

        def normalize_currency(x: str | None) -> str:
            if not x:
                return ""
            v = str(x).strip().upper()
            if v in {"USD", "US$", "$", "$US", "$USD"}:
                return "USD"
            if v in {"CHF", "SFR", "FR.", "FR", "CHF."}:
                return "CHF"
            if v in {"EUR", "EURO", "€"}:
                return "Euro"
            if v in {"CAD", "C$", "CA$"}:
                return "CAD"
            if v in {"USD", "CHF", "CAD"}:
                return v
            return ""  

        # VAT normalization: accept 8.1, 2.6, 3.8; if missing/invalid, default to 8.1
        def as_vat(x):
            try:
                v = float(x)
                if 0 < v < 1:
                    v = round(v * 100, 3)
                # Map close values
                if 7.9 <= v <= 8.2:
                    v = 8.1
                elif 2.4 <= v <= 2.7:
                    v = 2.6
                elif 3.6 <= v <= 3.9:
                    v = 3.8
                return v if v in {8.1, 2.6, 3.8} else 8.1
            except Exception:
                return 8.1

        result = {
            "amount": as_float(data.get("amount")),
            "attendees": (data.get("attendees") or "").strip(),
            "occasion": (data.get("occasion") or "").strip(),
            "payment_method": (data.get("payment_method") or "").strip(),
            "date": (data.get("date") or "").strip(),
            "category": normalized_category,
            "name": (data.get("name") or "").strip(),
            "vat_rate": as_vat(data.get("vat_rate")),
            "currency": normalize_currency(data.get("currency")),
        }

        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ===== Seeder: create sample expenses in Airtable =====
@app.post("/api/seed-samples")
def seed_samples():
    """Create N sample expenses in Airtable with fully filled, sensible values.
    Query params:
      - count: number of records (default 10, max 50)
      - month: YYYY-MM for the Date field (default current month)
      - statuses: optional comma-separated statuses to draw from

    Constraints for generated data:
      - Names (attendees/reimburse to) come from: Bob, Alice, Ayesha, Emily, Abtin, Niklas
      - VAT Rate: one of 8.1, 2.6, 3.8 (mapped sensibly by category/merchant)
      - Receipt URL: https://google.com (no real image attached)
      - Currency: CHF, Euro, USD, or CAD
    """
    try:
        # Count and month
        try:
            count = int(request.args.get("count") or 10)
        except Exception:
            count = 10
        count = max(1, min(count, 50))

        today = date.today()
        period = request.args.get("month") or f"{today.year:04d}-{today.month:02d}"

        # Status pool (validated)
        statuses_param = request.args.get("statuses")
        status_pool = [s.strip() for s in statuses_param.split(",")] if statuses_param else REPORT_INCLUDE_STATUSES
        status_pool = [s for s in status_pool if s in ALLOWED_STATUSES] or ["Under Review"]

        table = get_airtable_table()

        # Fixed pools for deterministic, sensible values
        categories_pool = ["Meals", "Travels", "Supplies", "Others"]
        payments_pool = ["Company card", "Personal", "Cash", "Other"]
        currencies_pool = ["CHF", "Euro", "USD", "CAD"]
        people_pool = ["Bob", "Alice", "Ayesha", "Emily", "Abtin", "Niklas"]

        merchants_by_cat = {
            "Meals": ["Starbucks", "Local Bakery", "Pizza Place", "Sushi Bar", "Cafeteria"],
            "Travels": ["Uber", "SBB", "Swiss Air", "Hotel City", "Taxi Basel"],
            "Supplies": ["Staples", "Office Depot", "Migros", "Coop", "Hardware Store"],
            "Others": ["Amazon", "Apple Store", "Fnac", "Kiosk", "General Store"],
        }

        occasions_by_cat = {
            "Meals": ["Team lunch", "Coffee with client", "Dinner with partner"],
            "Travels": ["Taxi to client", "Train to Zurich", "Flight to meeting", "Hotel night"],
            "Supplies": ["Office supplies", "Printer paper", "Notebook & pens", "Cables & adapters"],
            "Others": ["Subscription", "Small equipment", "Misc expense"],
        }

        def rand_day(period_str: str) -> str:
            try:
                y, m = period_str.split("-")
                y = int(y); m = int(m)
                import calendar
                last_day = calendar.monthrange(y, m)[1]
                d = random.randint(1, last_day)
                return f"{y:04d}-{m:02d}-{d:02d}"
            except Exception:
                return today.isoformat()

        def amount_for_category(cat: str) -> float:
            if cat == "Meals":
                return round(random.uniform(12, 85), 2)
            if cat == "Supplies":
                return round(random.uniform(9, 220), 2)
            if cat == "Travels":
                return round(random.uniform(15, 520), 2)
            return round(random.uniform(7, 160), 2)

        def vat_for(cat: str, merchant: str) -> float:
            m = merchant.lower()
            if cat == "Meals":
                return 8.1
            if cat == "Travels":
                # Accommodation uses 3.8
                if "hotel" in m:
                    return 3.8
                return 8.1
            if cat == "Supplies":
                # Groceries/books/water often 2.6; office/equipment might be 8.1 — choose sensibly
                if any(x in m for x in ["migros", "coop", "kiosk"]):
                    return 2.6
                return 8.1
            # Others — default 8.1
            return 8.1

        def choose_currency() -> str:
            # Slightly favor CHF and Euro
            roll = random.random()
            if roll < 0.45:
                return "CHF"
            if roll < 0.8:
                return "Euro"
            if roll < 0.92:
                return "USD"
            return "CAD"

        created = []
        for _ in range(count):
            cat = random.choice(categories_pool)
            merchant = random.choice(merchants_by_cat.get(cat, merchants_by_cat["Others"]))
            occ = random.choice(occasions_by_cat.get(cat, occasions_by_cat["Others"]))
            payer = random.choice(people_pool)
            paym = random.choice(payments_pool)
            curr = choose_currency()
            amt = amount_for_category(cat)
            day = rand_day(period)
            status = random.choice(status_pool)
            vat_rate = vat_for(cat, merchant)
            image_url = "https://google.com"

            attendees_list = random.sample(people_pool, k=random.randint(1, min(4, len(people_pool))))
            record_payload = {
                "Id": str(random.randint(1_000_000, 9_999_999)),
                "Name": f"{merchant} {occ}",
                "Amount": float(amt),
                "Attendees": ", ".join(attendees_list),
                "Occasion": occ,
                "Payment": paym,
                "Date": day,
                "Date added": today.isoformat(),
                "Category": cat,
                "Reimburse to": payer,
                "Status": status,
                "Currency": curr,
                "VAT Rate": vat_rate,
                "Receipt": [{"url": image_url}],
            }
            try:
                rec = table.create(record_payload, typecast=True)
                created.append(rec.get("id"))
            except Exception as e:
                return jsonify({"ok": False, "error": f"Airtable create failed: {e}", "created": created}), 502

        return jsonify({
            "ok": True,
            "count": len(created),
            "period": period,
            "recordIds": created,
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

if __name__ == "__main__":
    main()
