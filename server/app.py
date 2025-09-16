"""Flask backend for the In-House Expensify expense tracker."""
from __future__ import annotations

import base64
import hashlib
import json
import os
import random
import re
import smtplib
import sys
import time
from datetime import date, datetime
from email.message import EmailMessage
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable, Mapping

from dotenv import load_dotenv
from flask import Flask, Response, jsonify, request, send_from_directory
from flask_cors import CORS
from openai import OpenAI
from werkzeug.datastructures import FileStorage
from werkzeug.utils import secure_filename

load_dotenv()

ROOT_DIR = Path(__file__).resolve().parent.parent
PUBLIC_DIR = ROOT_DIR / "public"
UPLOAD_DIR = ROOT_DIR / (os.getenv("UPLOAD_DIR") or "uploads")
SERVER_DIR = Path(__file__).resolve().parent
if str(SERVER_DIR) not in sys.path:
    sys.path.insert(0, str(SERVER_DIR))
try:
    import reporting  # type: ignore
except Exception:
    reporting = None

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10 MB

allowed_origin = os.getenv("ALLOWED_ORIGIN", "*")
if allowed_origin == "*":
    CORS(app)
else:
    CORS(app, resources={r"/*": {"origins": [allowed_origin]}})

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

AIRTABLE_API_KEY = os.getenv("AIRTABLE_API_KEY")
AIRTABLE_BASE_ID = os.getenv("AIRTABLE_BASE_ID")
AIRTABLE_TABLE_NAME = os.getenv("AIRTABLE_TABLE_NAME", "Expenses")
AIRTABLE_TABLE_ID = os.getenv("AIRTABLE_TABLE_ID")
AIRTABLE_URL = os.getenv("AIRTABLE_URL")
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL")
DEFAULT_CURRENCY = os.getenv("DEFAULT_CURRENCY") or "Euro"
FX_RATES_CHF_JSON = os.getenv("FX_RATES_CHF_JSON")

ALLOWED_CATEGORIES = ["Travels", "Meals", "Supplies", "Others"]
ALLOWED_STATUSES = ["Done", "In-Progress", "Under Review"]
REPORT_INCLUDE_STATUSES = [
    s.strip()
    for s in (os.getenv("REPORT_INCLUDE_STATUSES") or "Done,In-Progress,Under Review").split(",")
    if s.strip()
]

SMTP_HOST = os.getenv("SMTP_HOST")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASS = os.getenv("SMTP_PASS")
EMAIL_FROM = os.getenv("EMAIL_FROM")
EMAIL_TO = os.getenv("EMAIL_TO")

STORAGE_PROVIDER = (os.getenv("STORAGE_PROVIDER") or "local").lower()
CLD_CLOUD_NAME = os.getenv("CLOUDINARY_CLOUD_NAME")
CLD_API_KEY = os.getenv("CLOUDINARY_API_KEY")
CLD_API_SECRET = os.getenv("CLOUDINARY_API_SECRET")
CLD_FOLDER = os.getenv("CLOUDINARY_FOLDER") or "in-house-expensify"


def _configure_cloudinary() -> bool:
    """Configure Cloudinary if requested by the environment."""
    if STORAGE_PROVIDER != "cloudinary":
        return False
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
    except Exception:
        return False
    return True


# Structural change: configure Cloudinary once at import time.
CLOUDINARY_ENABLED = _configure_cloudinary()


def _parse_airtable_url(url: str | None) -> tuple[str | None, str | None]:
    """Extract base (app...) and table (tbl...) IDs from an Airtable UI URL."""
    if not url:
        return None, None
    base_match = re.search(r"(app[a-zA-Z0-9]+)", url)
    table_match = re.search(r"(tbl[a-zA-Z0-9]+)", url)
    base_id = base_match.group(1) if base_match else None
    table_id = table_match.group(1) if table_match else None
    return base_id, table_id


def get_airtable_table() -> Any:
    """Return a pyairtable Table configured from environment variables."""
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
        raise RuntimeError(
            "Missing Airtable table. Set AIRTABLE_TABLE_ID, AIRTABLE_TABLE_NAME, or AIRTABLE_URL."
        )

    return Table(AIRTABLE_API_KEY, base_id, table_segment)


# Structural change: consolidated period parsing shared by multiple routes.
def _current_period() -> str:
    """Return the current year-month period in YYYY-MM format."""
    today = date.today()
    return f"{today.year:04d}-{today.month:02d}"


def _period_from_request(*, allow_body: bool = False, allow_parts: bool = False) -> str:
    """Resolve a YYYY-MM period from request args or body with sensible defaults."""
    period = request.args.get("month")
    if not period and allow_body:
        payload: Mapping[str, Any] | None = None
        if request.is_json:
            candidate = request.get_json(silent=True) or {}
            if isinstance(candidate, Mapping):
                payload = candidate
        if payload is None:
            payload = request.form
        raw_period = payload.get("month") if payload else None
        if isinstance(raw_period, str):
            period = raw_period
    if not period and allow_parts:
        year = (request.args.get("year") or str(date.today().year)).zfill(4)
        month_value = (
            request.args.get("monthNum")
            or request.args.get("m")
            or str(date.today().month)
        )
        period = f"{year}-{str(month_value).zfill(2)}"
    return period or _current_period()


# Structural change: consolidated status parsing and default handling.
def _parse_statuses_param(
    raw: str | None,
    *,
    allowed: Iterable[str] | None = None,
    default: Iterable[str] | None = None,
) -> list[str] | None:
    """Normalize a comma separated statuses parameter."""
    statuses: list[str] = []
    if raw:
        statuses = [s.strip() for s in raw.split(",") if s.strip()]
        if allowed is not None:
            statuses = [s for s in statuses if s in allowed]
        if statuses:
            return statuses
    if default is not None:
        filtered_default = [s for s in default if allowed is None or s in allowed]
        return filtered_default or None
    return None


# Structural change: pulled VAT normalization into a shared helper.
def _normalize_vat_rate(raw: Any) -> float | None:
    """Return a VAT rate of 8.1, 2.6 or 3.8 if possible."""
    if raw in (None, ""):
        return None
    try:
        value = float(str(raw).strip())
    except (TypeError, ValueError):
        return None
    if 0 < value < 1:
        value = round(value * 100, 3)
    allowed_rates = {8.1, 2.6, 3.8}
    return value if value in allowed_rates else None


# Structural change: pulled currency normalization into a shared helper.
def _normalize_currency(raw: str | None) -> str | None:
    """Map a currency value or symbol to an allowed currency name."""
    if not raw:
        return None
    value = str(raw).strip().upper()
    mapping = {
        "USD": "USD",
        "US$": "USD",
        "$": "USD",
        "$US": "USD",
        "$USD": "USD",
        "CHF": "CHF",
        "SFR": "CHF",
        "FR.": "CHF",
        "FR": "CHF",
        "CHF.": "CHF",
        "EUR": "Euro",
        "EURO": "Euro",
        "€": "Euro",
        "CAD": "CAD",
        "C$": "CAD",
        "CA$": "CAD",
    }
    if value in mapping:
        return mapping[value]
    if value in {"USD", "CHF", "CAD"}:
        return value
    return None


def _parse_amount(raw: str | None) -> float:
    """Convert an amount form value to float with clear error messages."""
    if raw is None or str(raw).strip() == "":
        raise ValueError("Amount is required")
    try:
        return float(raw)
    except ValueError as exc:  # pragma: no cover - defensive guard
        raise ValueError("Amount must be a number") from exc


# Structural change: unified receipt storage for local and Cloudinary providers.
def _store_receipt(image: FileStorage, img_bytes: bytes) -> tuple[str | None, list[dict[str, Any]] | None]:
    """Persist the uploaded receipt and return the public URL and Airtable attachment payload."""
    if STORAGE_PROVIDER == "cloudinary" and CLOUDINARY_ENABLED:
        try:
            from cloudinary import uploader as cld_uploader

            upload_res = cld_uploader.upload(
                BytesIO(img_bytes),
                folder=CLD_FOLDER,
                resource_type="image",
            )
        except Exception as exc:  # pragma: no cover - depends on external service
            raise RuntimeError(f"Cloudinary upload failed: {exc}") from exc
        image_url = upload_res.get("secure_url") or upload_res.get("url")
        return image_url, ([{"url": image_url}] if image_url else None)

    filename = secure_filename(image.filename or "receipt")
    name_root, ext = os.path.splitext(filename)
    safe_root = name_root.replace(" ", "_") or "receipt"
    final_name = f"{int(time.time() * 1000)}_{safe_root}{ext}"
    save_path = UPLOAD_DIR / final_name
    with open(save_path, "wb") as fh:
        fh.write(img_bytes)
    image_url = build_public_url(f"/uploads/{final_name}")
    return image_url, ([{"url": image_url}])


def _fx_rates_chf() -> dict[str, float]:
    """Return configured FX rates to CHF or sensible defaults."""
    try:
        if FX_RATES_CHF_JSON:
            data = json.loads(FX_RATES_CHF_JSON)
            if isinstance(data, dict):
                return {k.upper(): float(v) for k, v in data.items()}
    except Exception:  # pragma: no cover - permissive parsing
        pass
    return {
        "CHF": 1.0,
        "EUR": 0.96,
        "EURO": 0.96,
        "USD": 0.90,
        "CAD": 0.66,
    }


def _fx_policy_description(rates: dict[str, float]) -> str:
    """Return a human readable description of the FX policy."""
    try:
        parts = [f"1 {cur} = {float(rate):.4f} CHF" for cur, rate in sorted(rates.items())]
        if not parts:
            return "Standard 1:1 CHF conversion"
        return "Internal fixed conversion policy: " + ", ".join(parts)
    except Exception:  # pragma: no cover - defensive guard
        return "Standard 1:1 CHF conversion"


def _to_chf(amount: float, currency: str | None) -> float:
    """Convert an amount to CHF using configured FX rates."""
    try:
        cur = (currency or "CHF").upper()
        rate = _fx_rates_chf().get(cur, 1.0)
        return round(float(amount) * float(rate), 2)
    except Exception:  # pragma: no cover - defensive guard
        return round(float(amount or 0), 2)


def build_public_url(path_segment: str) -> str:
    """Construct a fully qualified public URL for a stored asset."""
    base = (PUBLIC_BASE_URL or request.host_url).rstrip("/")
    if not path_segment.startswith("/"):
        path_segment = "/" + path_segment
    return f"{base}{path_segment}"


def _b64url_decode(data: str) -> bytes:
    """Decode base64url strings from Cloudflare Access headers."""
    padded = data + "=" * (-len(data) % 4)
    return base64.urlsafe_b64decode(padded)


def _cf_email_from_request() -> str | None:
    """Extract user email from Cloudflare Access headers or cookies."""
    email = request.headers.get("Cf-Access-Authenticated-User-Email")
    if email:
        return email
    xf_user = request.headers.get("X-Forwarded-User")
    if xf_user and "@" in xf_user:
        return xf_user
    token = request.cookies.get("CF_Authorization") or request.headers.get("Cf-Access-Jwt-Assertion")
    if token and token.count(".") >= 2:
        try:
            payload_raw = _b64url_decode(token.split(".")[1])
            data = json.loads(payload_raw.decode("utf-8"))
            em = data.get("email") or data.get("sub")
            if isinstance(em, str) and "@" in em:
                return em
        except Exception:  # pragma: no cover - permissive parsing
            return None
    return None


@app.route("/version")
def version() -> dict[str, str]:
    """Return the backend version."""
    return {"version": "1.3.0"}


@app.post("/api/expenses")
def create_expense():
    """Create an expense, upload the receipt and sync it to Airtable."""
    try:
        image = request.files.get("image")
        if image is None or image.filename == "":
            return jsonify({"error": "Image is required"}), 400

        img_bytes = image.read()
        if not img_bytes:
            return jsonify({"error": "Empty image"}), 400

        img_hash = hashlib.sha256(img_bytes).hexdigest()
        image.stream = BytesIO(img_bytes)

        try:
            amount = _parse_amount(request.form.get("amount"))
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400

        attendees = request.form.get("attendees", "")
        occasion = request.form.get("occasion", "")
        payment_method = request.form.get("payment_method", "")
        date_str = request.form.get("date") or date.today().isoformat()
        category = (request.form.get("category") or "Uncategorized").strip()
        reimburse_to = (request.form.get("reimburse_to") or "None").strip() or "None"

        vat_rate = _normalize_vat_rate(request.form.get("vat_rate"))
        currency = _normalize_currency(request.form.get("currency")) or DEFAULT_CURRENCY

        name_val = (request.form.get("name") or "").strip()
        if not name_val:
            name_val = f"Expense {date_str}"
        random_id = str(random.randint(1_000_000, 9_999_999))

        try:
            image_url, attachment = _store_receipt(image, img_bytes)
        except RuntimeError as exc:
            return jsonify({"error": str(exc)}), 502

        table = get_airtable_table()
        uploader_email = _cf_email_from_request()

        original_amount = amount
        original_currency = currency
        amount_chf = _to_chf(original_amount, original_currency)

        payload: dict[str, Any] = {
            "Id": random_id,
            "Name": name_val,
            "Amount": amount_chf,
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
            "Currency": original_currency,
            "Original Amount": original_amount,
        }
        if vat_rate is not None:
            payload["VAT Rate"] = vat_rate
        if uploader_email:
            payload["Uploaded By"] = uploader_email

        try:
            record = table.create(payload, typecast=True)
        except Exception as exc:
            return (
                jsonify(
                    {
                        "error": f"Airtable create failed: {exc}",
                        "hint": "Check AIRTABLE_URL or AIRTABLE_BASE_ID and table name/id match your base.",
                    }
                ),
                502,
            )

        return jsonify(
            {
                "ok": True,
                "recordId": record.get("id"),
                "imageUrl": image_url,
                "uploadedBy": uploader_email,
            }
        )
    except Exception as exc:  # pragma: no cover - defensive guard
        return jsonify({"error": str(exc)}), 500


@app.get("/uploads/<path:filename>")
def uploads(filename: str):
    """Serve uploaded receipt files."""
    return send_from_directory(UPLOAD_DIR, filename)


@app.get("/")
def index():
    """Serve the application index page."""
    return send_from_directory(PUBLIC_DIR, "index.html")


@app.get("/styles.css")
def styles():
    """Serve the compiled stylesheet."""
    return send_from_directory(PUBLIC_DIR, "styles.css")


@app.get("/app.js")
def app_js():
    """Serve the application bundle."""
    return send_from_directory(PUBLIC_DIR, "app.js")


@app.get("/records")
def records_page():
    """Serve the records page."""
    return send_from_directory(PUBLIC_DIR, "records.html")


@app.get("/report")
def report_page():
    """Serve the report page."""
    return send_from_directory(PUBLIC_DIR, "report.html")


@app.get("/healthz")
def healthz() -> dict[str, str]:
    """Liveness probe used by deployment platforms."""
    return {"status": "ok"}


@app.errorhandler(404)
def handle_404(error):  # type: ignore[override]
    """Return JSON for missing API routes while preserving Flask defaults elsewhere."""
    if request.path.startswith("/api/"):
        return jsonify({"error": "Not found", "path": request.path}), 404
    return error


@app.errorhandler(405)
def handle_405(error):  # type: ignore[override]
    """Return JSON for invalid method calls on API routes."""
    if request.path.startswith("/api/"):
        return jsonify({"error": "Method not allowed", "path": request.path}), 405
    return error


def main() -> None:
    """Run the Flask development server."""
    port = int(os.getenv("PORT", "3000"))
    app.run(host="0.0.0.0", port=port, debug=os.getenv("FLASK_DEBUG") == "1")


@app.get("/api/expenses")
def list_expenses():
    """Return all expenses for a month including duplicate detection hints."""
    try:
        month = _period_from_request()
        formula = f"DATETIME_FORMAT({{Date added}}, 'YYYY-MM') = '{month}'"
        table = get_airtable_table()
        records = table.all(formula=formula)

        items = []
        for record in records:
            fields = record.get("fields", {})
            legacy_approved = fields.get("Approved")
            status_val = fields.get("Status")
            if not status_val and legacy_approved is not None:
                status_val = "Done" if legacy_approved else "Under Review"

            items.append(
                {
                    "record_id": record.get("id"),
                    "id": fields.get("id"),
                    "name": fields.get("Name"),
                    "amount": fields.get("Amount"),
                    "attendees": fields.get("Attendees"),
                    "occasion": fields.get("Occasion"),
                    "payment": fields.get("Payment"),
                    "date": fields.get("Date"),
                    "date_added": fields.get("Date added"),
                    "category": fields.get("Category"),
                    "reimburse_to": fields.get("Reimburse to"),
                    "currency": fields.get("Currency"),
                    "original_amount": fields.get("Original Amount"),
                    "original_currency": fields.get("Currency"),
                    "vat_rate": fields.get("VAT Rate"),
                    "status": status_val or "Under Review",
                    "receipt_url": _first_url(fields.get("Receipt")),
                    "hash": fields.get("Hash"),
                    "uploaded_by": fields.get("Uploaded By"),
                }
            )

        hash_counts: dict[str, int] = {}
        for item in items:
            file_hash = item.get("hash")
            if file_hash:
                hash_counts[file_hash] = hash_counts.get(file_hash, 0) + 1
        for item in items:
            file_hash = item.get("hash")
            count = hash_counts.get(file_hash or "", 0)
            item["duplicate_hint"] = bool(file_hash and count > 1)
            item["duplicate_count"] = count if file_hash else 0

        items.sort(key=lambda x: (x.get("date_added") or ""), reverse=True)
        return jsonify({"month": month, "count": len(items), "items": items})
    except Exception as exc:  # pragma: no cover - defensive guard
        return jsonify({"error": str(exc)}), 500


@app.post("/api/expenses/<record_id>/status")
def update_status(record_id: str):
    """Update the status for a specific expense record."""
    try:
        data = request.get_json(silent=True) or request.form
        status = (data.get("status") or "").strip()
        if status not in ALLOWED_STATUSES:
            return (
                jsonify({"ok": False, "error": "Invalid status", "allowed": ALLOWED_STATUSES}),
                400,
            )
        table = get_airtable_table()
        updated = table.update(record_id, {"Status": status}, typecast=True)
        return jsonify({"ok": True, "recordId": record_id, "status": updated.get("fields", {}).get("Status")})
    except Exception as exc:  # pragma: no cover - defensive guard
        return jsonify({"ok": False, "error": str(exc)}), 500


def _email_recipients() -> list[str]:
    """Return configured email recipients if EMAIL_TO is set."""
    if not EMAIL_TO:
        return []
    return [addr.strip() for addr in EMAIL_TO.split(",") if addr.strip()]


def send_email(
    subject: str,
    html_body: str,
    to_addrs: list[str],
    attachments: list[dict[str, Any]] | None = None,
) -> None:
    """Send an HTML email using the configured SMTP server."""
    if not (SMTP_HOST and EMAIL_FROM and to_addrs):
        return

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = EMAIL_FROM
    msg["To"] = ", ".join(to_addrs)
    msg.set_content("This email requires an HTML-capable client.")
    msg.add_alternative(html_body, subtype="html")

    if attachments:
        for attachment in attachments:
            if not attachment:
                continue
            filename = attachment.get("filename") or "attachment"
            content = attachment.get("content")
            maintype = attachment.get("maintype") or "application"
            subtype = attachment.get("subtype") or "octet-stream"
            if isinstance(content, str):
                content = content.encode("utf-8")
            if content is None:
                continue
            msg.add_attachment(content, maintype=maintype, subtype=subtype, filename=filename)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=12) as server:
        try:
            server.starttls()
        except Exception:  # pragma: no cover - some servers disallow STARTTLS
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
) -> None:
    """Send a notification email for a new expense if SMTP is configured."""
    recipients = _email_recipients()
    if not recipients:
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
      <li><strong>Attendees:</strong> {attendees or '(none)'} </li>
      <li><strong>Receipt:</strong> <a href="{image_url}">View</a></li>
      {f'<li><strong>Airtable Record ID:</strong> {airtable_record_id}</li>' if airtable_record_id else ''}
    </ul>
    """
    send_email(subject, body, recipients)


@app.get("/api/test-email")
def test_email():
    """Send a test email to EMAIL_TO to verify SMTP configuration."""
    recipients = _email_recipients()
    if not SMTP_HOST or not EMAIL_FROM:
        return (
            jsonify(
                {
                    "ok": False,
                    "error": "SMTP not configured",
                    "hint": "Set SMTP_HOST, EMAIL_FROM, and (optionally) SMTP_PORT/SMTP_USER/SMTP_PASS in .env",
                }
            ),
            400,
        )
    if not recipients:
        return (
            jsonify(
                {
                    "ok": False,
                    "error": "No recipients configured",
                    "hint": "Set EMAIL_TO to one or more emails (comma-separated) in .env",
                }
            ),
            400,
        )
    try:
        body = """
        <p>This is a test email from the In-House Expensify app.</p>
        <p>If you received this, SMTP is configured correctly.</p>
        """
        send_email("In-House Expensify — Test Email", body, recipients)
        return jsonify({"ok": True})
    except Exception as exc:  # pragma: no cover - depends on SMTP server
        return jsonify({"ok": False, "error": str(exc)}), 500


def _normalize_status(status: str | None) -> str:
    """Normalize status strings from Airtable to the allowed set."""
    value = (status or "").strip()
    lower = value.lower().replace("_", "-").replace(" ", "-")
    if lower in {"done"}:
        return "Done"
    if lower in {"in-progress", "inprogress"}:
        return "In-Progress"
    if lower in {"under-review", "underreview", "under-review."}:
        return "Under Review"
    if value in {"True", "true", "1"}:
        return "Done"
    if value in {"False", "false", "0"}:
        return "Under Review"
    return value or "Under Review"


def _normalize_payment_method(payment: str | None) -> str:
    """Normalize payment method labels."""
    value = (payment or "").strip().lower()
    if not value:
        return "Other"
    if value in {"company card", "company-card", "companycard"}:
        return "Company card"
    if value in {"personal", "personal reimbursement", "personal-reimbursement", "reimbursement"}:
        return "Personal"
    if value == "cash":
        return "Cash"
    return payment.strip() if payment else "Other"


def _first_url(attachments: Any) -> str | None:
    """Return the first attachment URL if present."""
    try:
        if isinstance(attachments, list) and attachments:
            return attachments[0].get("url")
    except Exception:  # pragma: no cover - defensive guard
        return None
    return None


def _round2(value: float | int | None) -> float:
    """Round values to two decimals without raising."""
    try:
        return round(float(value or 0), 2)
    except Exception:  # pragma: no cover - defensive guard
        return 0.0


def _vat_amount(gross: float, vat_rate: float | None) -> float:
    """Compute the VAT amount for a gross value."""
    try:
        rate = float(vat_rate or 0)
        if 0 < rate < 1:
            rate *= 100
        return round(float(gross or 0) * (rate / 100.0), 2)
    except Exception:  # pragma: no cover - defensive guard
        return 0.0


def _build_monthly_report(period: str, *, include_statuses: list[str] | None = None) -> dict:
    """Build monthly report aggregations filtered by the given statuses."""
    include_set = set(include_statuses or REPORT_INCLUDE_STATUSES or ["Done"])
    formula = (
        f"IF({{Date}}, DATETIME_FORMAT({{Date}}, 'YYYY-MM') = '{period}', "
        f"DATETIME_FORMAT({{Date added}}, 'YYYY-MM') = '{period}')"
    )
    table = get_airtable_table()
    records = table.all(formula=formula)

    currency_buckets: dict[str, dict[str, Any]] = {}
    rows: list[dict[str, Any]] = []
    fx_rates = _fx_rates_chf()

    for record in records:
        fields = record.get("fields", {})
        amount = fields.get("Amount")
        currency = fields.get("Currency")
        if amount in (None, ""):
            continue
        try:
            gross = float(amount)
        except Exception:
            continue

        cur = str(currency).strip() if currency not in (None, "") else "Unknown"
        status = _normalize_status(fields.get("Status") if fields.get("Status") is not None else fields.get("Approved"))
        payment = _normalize_payment_method(fields.get("Payment"))
        payer = (fields.get("Payer") or fields.get("Reimburse to") or "").strip() or "Unknown"
        category = (fields.get("Category") or "").strip() or "Others"
        date_occ = fields.get("Date") or fields.get("Date added") or ""
        vat_rate = fields.get("VAT Rate")

        vat_amt = _vat_amount(gross, vat_rate)
        net_amt = round(gross - vat_amt, 2)

        bucket = currency_buckets.setdefault(
            cur,
            {
                "totals": {"gross": 0.0, "net": 0.0, "vat": 0.0},
                "byCategory": {},
                "byPaymentMethod": {},
                "companyCardCharged": 0.0,
                "reimbursementsByEmployee": {},
                "pending": {
                    "inProgress": {"count": 0, "gross": 0.0},
                    "underReview": {"count": 0, "gross": 0.0},
                },
            },
        )

        if status == "In-Progress":
            bucket["pending"]["inProgress"]["count"] += 1
            bucket["pending"]["inProgress"]["gross"] += gross
        elif status == "Under Review":
            bucket["pending"]["underReview"]["count"] += 1
            bucket["pending"]["underReview"]["gross"] += gross

        if status not in include_set:
            continue

        bucket["totals"]["gross"] += gross
        bucket["totals"]["net"] += net_amt
        bucket["totals"]["vat"] += vat_amt

        bucket["byCategory"][category] = bucket["byCategory"].get(category, 0.0) + gross
        bucket["byPaymentMethod"][payment] = bucket["byPaymentMethod"].get(payment, 0.0) + gross

        if payment.lower().startswith("company"):
            bucket["companyCardCharged"] += gross
        if payment.lower() in {"personal", "cash"}:
            bucket["reimbursementsByEmployee"][payer] = bucket["reimbursementsByEmployee"].get(payer, 0.0) + gross

        rows.append(
            {
                "recordId": record.get("id"),
                "date": date_occ,
                "payer": payer,
                "category": category,
                "paymentMethod": payment,
                "gross": gross,
                "net": net_amt,
                "vat": vat_amt,
                "currency": cur,
                "status": status,
                "fxPolicy": _fx_policy_description(fx_rates),
            }
        )

    return {
        "period": period,
        "rows": rows,
        "currencyBuckets": currency_buckets,
        "fxPolicy": _fx_policy_description(fx_rates),
    }


def generate_report_pdf(period: str, *, include_statuses: list[str] | None = None) -> tuple[bytes, dict]:
    """Build the monthly report and render it to PDF using the reporting module."""
    report = _build_monthly_report(period, include_statuses=include_statuses)
    if reporting is None:
        raise RuntimeError("reporting module not available")
    pdf_bytes = reporting.render_report_pdf(report)
    return pdf_bytes, report


def _render_error_pdf(message: str) -> bytes:
    """Render a tiny PDF that displays an error message."""
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib.units import mm
        from reportlab.platypus import Paragraph, SimpleDocTemplate
    except Exception:
        return (message or "Error").encode("utf-8")
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=16 * mm, bottomMargin=16 * mm)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("Report Generation Error", styles["Heading2"]),
        Paragraph(message or "Unknown error", styles["BodyText"]),
    ]
    doc.build(elements)
    return buffer.getvalue()


@app.get("/api/expense-report")
def get_expense_report():
    """Return the monthly expense report as JSON."""
    try:
        period = _period_from_request(allow_parts=True)
        include_statuses = _parse_statuses_param(request.args.get("statuses"))
        report = _build_monthly_report(period, include_statuses=include_statuses)
        return jsonify(report)
    except Exception as exc:  # pragma: no cover - defensive guard
        return jsonify({"error": str(exc)}), 500


@app.get("/api/expense-report.check")
def expense_report_check():
    """Return diagnostics for report generation dependencies."""
    checks: dict[str, Any] = {}
    try:
        import reportlab  # type: ignore

        checks["reportlab"] = getattr(reportlab, "__version__", "present")
    except Exception as exc:
        checks["reportlab"] = f"missing: {exc}"
    try:
        table = get_airtable_table()
        records = table.all(max_records=1)
        checks["airtable"] = f"ok ({len(records)} accessible)"
    except Exception as exc:  # pragma: no cover - depends on Airtable
        checks["airtable"] = f"error: {exc}"
    return jsonify(checks)


@app.post("/api/send-summary")
def send_summary_email():
    """Generate the monthly PDF report and email it to the authenticated user."""
    period = _period_from_request(allow_body=True)
    try:
        statuses_param = request.args.get("statuses")
        if not statuses_param and request.is_json:
            payload = request.get_json(silent=True) or {}
            if isinstance(payload, Mapping):
                statuses_param = payload.get("statuses")
        if not statuses_param:
            statuses_param = request.form.get("statuses")
        include_statuses = _parse_statuses_param(statuses_param)
        pdf_bytes, report = generate_report_pdf(period, include_statuses=include_statuses)
        timestamp = datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
        html = f"""
        <p>Please find the attached PDF summary for {report['period']}.<br/>
        Generated at {timestamp}.</p>
        """
    except Exception as exc:
        msg = str(exc)
        if "reportlab" in msg.lower():
            return jsonify({"ok": False, "error": f"PDF render failed: {exc}"}), 500
        return jsonify({"ok": False, "error": f"Airtable error: {exc}"}), 502

    attachments = [
        {
            "filename": f"expenses_{report['period']}.pdf",
            "content": pdf_bytes,
            "maintype": "application",
            "subtype": "pdf",
        }
    ]

    if not (SMTP_HOST and EMAIL_FROM):
        return jsonify({"ok": False, "error": "SMTP not configured (SMTP_HOST/EMAIL_FROM)"}), 400

    current_email = _cf_email_from_request()
    if not current_email:
        return jsonify({"ok": False, "error": "No authenticated user email from SSO"}), 401

    recipients = [current_email]
    subject = f"Expense Report {report['period']} — rows: {len(report.get('rows', []))}"

    try:
        import threading

        threading.Thread(
            target=lambda: send_email(subject, html, recipients, attachments=attachments),
            daemon=True,
        ).start()
    except Exception:  # pragma: no cover - fallback execution path
        try:
            send_email(subject, html, recipients, attachments=attachments)
        except Exception as exc:
            return jsonify({"ok": False, "error": str(exc)}), 500

    total = sum((row.get("gross") or 0) for row in report.get("rows", []))
    return jsonify(
        {
            "ok": True,
            "month": report["period"],
            "count": len(report.get("rows", [])),
            "total": _round2(total),
            "queued": True,
        }
    )


@app.get("/api/expense-report.pdf")
def get_expense_report_pdf():
    """Return the monthly expense report PDF for download."""
    try:
        period = _period_from_request(allow_parts=True)
        include_statuses = _parse_statuses_param(request.args.get("statuses"))
        pdf_bytes, report = generate_report_pdf(period, include_statuses=include_statuses)
        response = Response(pdf_bytes, mimetype="application/pdf")
        response.headers["Content-Disposition"] = f"inline; filename=expenses_{report['period']}.pdf"
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        return response
    except Exception as exc:
        pdf_bytes = reporting.render_error_pdf(str(exc)) if reporting is not None else _render_error_pdf(str(exc))
        response = Response(pdf_bytes, mimetype="application/pdf", status=500)
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        return response


@app.get("/api/whoami")
def whoami():
    """Return identity information from upstream SSO headers."""
    email = _cf_email_from_request()
    user = email or request.headers.get("X-Forwarded-User") or ""
    source = "cloudflare-access" if email else "unauthenticated"
    return jsonify(
        {
            "email": email or user,
            "source": source,
            "headersPresent": [
                header
                for header in ["Cf-Access-Authenticated-User-Email", "X-Forwarded-User"]
                if request.headers.get(header)
            ],
        }
    )


def _openai_client() -> OpenAI:
    """Return an OpenAI client configured from the environment."""
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("Missing OPENAI_API_KEY")
    return OpenAI(api_key=api_key)


@app.post("/api/analyze")
def analyze_receipt():
    """Send a receipt image to OpenAI for extraction and normalize the response."""
    try:
        if "image" not in request.files or request.files["image"].filename == "":
            return jsonify({"error": "Image is required"}), 400
        uploaded = request.files["image"]
        img_bytes = uploaded.read()
        mime = uploaded.mimetype or "image/jpeg"
        data_url = f"data:{mime};base64,{base64.b64encode(img_bytes).decode('ascii')}"

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
        response = client.chat.completions.create(
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

        content = response.choices[0].message.content or "{}"
        try:
            data = json.loads(content)
        except Exception:
            data = {}

        def as_float(value: Any) -> float:
            try:
                return float(value)
            except Exception:
                return 0.0

        def normalize_category(raw_value: str) -> str:
            value = (raw_value or "").strip().lower()
            if not value:
                return ""
            travel_syn = ["travel", "travels", "trip", "transport", "transportation", "flight", "train", "taxi", "uber", "lyft", "car"]
            meals_syn = ["meal", "meals", "food", "lunch", "dinner", "breakfast", "restaurant", "cafe", "coffee"]
            supplies_syn = ["supply", "supplies", "office supplies", "stationery", "hardware", "equipment"]
            if value in travel_syn or any(word in value for word in travel_syn):
                return "Travels"
            if value in meals_syn or any(word in value for word in meals_syn):
                return "Meals"
            if value in supplies_syn or any(word in value for word in supplies_syn):
                return "Supplies"
            for allowed in ALLOWED_CATEGORIES:
                if value == allowed.lower():
                    return allowed
            return "Others"

        normalized_category = normalize_category((data.get("category") or "").strip())
        normalized_currency = _normalize_currency(data.get("currency")) or ""
        normalized_vat = _normalize_vat_rate(data.get("vat_rate"))
        if normalized_vat is None:
            normalized_vat = 8.1

        result = {
            "amount": as_float(data.get("amount")),
            "attendees": (data.get("attendees") or "").strip(),
            "occasion": (data.get("occasion") or "").strip(),
            "payment_method": (data.get("payment_method") or "").strip(),
            "date": (data.get("date") or "").strip(),
            "category": normalized_category,
            "name": (data.get("name") or "").strip(),
            "vat_rate": normalized_vat,
            "currency": normalized_currency,
        }

        return jsonify(result)
    except Exception as exc:  # pragma: no cover - defensive guard
        return jsonify({"error": str(exc)}), 500


@app.post("/api/seed-samples")
def seed_samples():
    """Create sample expenses in Airtable for demonstration purposes."""
    try:
        try:
            count = int(request.args.get("count") or 10)
        except Exception:
            count = 10
        count = max(1, min(count, 50))

        period = request.args.get("month") or _current_period()
        status_pool = _parse_statuses_param(
            request.args.get("statuses"),
            allowed=ALLOWED_STATUSES,
            default=REPORT_INCLUDE_STATUSES,
        ) or ["Under Review"]

        table = get_airtable_table()

        categories_pool = ["Meals", "Travels", "Supplies", "Others"]
        payments_pool = ["Company card", "Personal", "Cash", "Other"]
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

        today = date.today()

        def rand_day(period_str: str) -> str:
            try:
                year_str, month_str = period_str.split("-")
                year_val = int(year_str)
                month_val = int(month_str)
                import calendar

                last_day = calendar.monthrange(year_val, month_val)[1]
                day_val = random.randint(1, last_day)
                return f"{year_val:04d}-{month_val:02d}-{day_val:02d}"
            except Exception:
                return today.isoformat()

        def amount_for_category(category: str) -> float:
            if category == "Meals":
                return round(random.uniform(12, 85), 2)
            if category == "Supplies":
                return round(random.uniform(9, 220), 2)
            if category == "Travels":
                return round(random.uniform(15, 520), 2)
            return round(random.uniform(7, 160), 2)

        def vat_for(category: str, merchant: str) -> float:
            merchant_lower = merchant.lower()
            if category == "Meals":
                return 8.1
            if category == "Travels":
                if "hotel" in merchant_lower:
                    return 3.8
                return 8.1
            if category == "Supplies":
                if any(keyword in merchant_lower for keyword in ["migros", "coop", "kiosk"]):
                    return 2.6
                return 8.1
            return 8.1

        def choose_currency() -> str:
            roll = random.random()
            if roll < 0.95:
                return "CHF"
            branch = random.random()
            if branch < 0.6:
                return "Euro"
            if branch < 0.9:
                return "USD"
            return "CAD"

        created: list[str | None] = []
        for _ in range(count):
            category = random.choice(categories_pool)
            merchant = random.choice(merchants_by_cat.get(category, merchants_by_cat["Others"]))
            occasion = random.choice(occasions_by_cat.get(category, occasions_by_cat["Others"]))
            payer = random.choice(people_pool)
            payment = random.choice(payments_pool)
            currency = choose_currency()
            amount_original = amount_for_category(category)
            amount_chf = _to_chf(amount_original, currency)
            day = rand_day(period)
            status = random.choice(status_pool)
            vat_rate = vat_for(category, merchant)
            attendees_list = random.sample(people_pool, k=random.randint(1, min(4, len(people_pool))))
            record_payload = {
                "Id": str(random.randint(1_000_000, 9_999_999)),
                "Name": f"{merchant} {occasion}",
                "Amount": float(amount_chf),
                "Attendees": ", ".join(attendees_list),
                "Occasion": occasion,
                "Payment": payment,
                "Date": day,
                "Date added": today.isoformat(),
                "Category": category,
                "Reimburse to": payer,
                "Status": status,
                "Currency": currency,
                "Original Amount": float(amount_original),
                "VAT Rate": vat_rate,
                "Receipt": [{"url": "https://google.com"}],
            }
            try:
                record = table.create(record_payload, typecast=True)
                created.append(record.get("id"))
            except Exception as exc:  # pragma: no cover - depends on Airtable
                return jsonify({"ok": False, "error": f"Airtable create failed: {exc}", "created": created}), 502

        return jsonify({"ok": True, "count": len(created), "period": period, "recordIds": created})
    except Exception as exc:  # pragma: no cover - defensive guard
        return jsonify({"ok": False, "error": str(exc)}), 500


if __name__ == "__main__":
    main()
