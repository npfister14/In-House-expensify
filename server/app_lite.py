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
import re

load_dotenv()

# Paths
ROOT_DIR = Path(__file__).resolve().parent.parent
PUBLIC_DIR = ROOT_DIR / "public"
UPLOAD_DIR = ROOT_DIR / (os.getenv("UPLOAD_DIR") or "uploads")

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
AIRTABLE_TABLE_ID = os.getenv("AIRTABLE_TABLE_ID")
AIRTABLE_URL = os.getenv("AIRTABLE_URL")
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL")
DEFAULT_CURRENCY = os.getenv("DEFAULT_CURRENCY") or "Euro"

ALLOWED_STATUSES = ["Done", "In-Progress", "Under Review"]


def _parse_airtable_url(url: str):
    if not url:
        return None, None
    base_match = re.search(r"(app[a-zA-Z0-9]+)", url)
    table_match = re.search(r"(tbl[a-zA-Z0-9]+)", url)
    return base_match.group(1) if base_match else None, table_match.group(1) if table_match else None


def get_airtable_table():
    if not AIRTABLE_API_KEY:
        raise RuntimeError("Missing AIRTABLE_API_KEY")
    from pyairtable import Table
    parsed_base, parsed_tbl = (None, None)
    if AIRTABLE_URL:
        parsed_base, parsed_tbl = _parse_airtable_url(AIRTABLE_URL)
    base_id = AIRTABLE_BASE_ID or parsed_base
    if not base_id:
        raise RuntimeError("Missing Airtable base. Set AIRTABLE_BASE_ID or AIRTABLE_URL.")
    table_segment = AIRTABLE_TABLE_ID or parsed_tbl or AIRTABLE_TABLE_NAME
    return Table(AIRTABLE_API_KEY, base_id, table_segment)


def build_public_url(path_segment: str) -> str:
    base = (PUBLIC_BASE_URL or request.host_url).rstrip('/')
    if not path_segment.startswith('/'):
        path_segment = '/' + path_segment
    return f"{base}{path_segment}"


@app.get("/healthz")
def healthz():
    return {"status": "ok"}

@app.get("/version")
def version():
    return {"version": "2.0.0"}

@app.get("/")
def index():
    return send_from_directory(PUBLIC_DIR, "index.html")


@app.get("/styles.css")
def styles():
    return send_from_directory(PUBLIC_DIR, "styles.css")


@app.get("/records")
def records_page():
    return send_from_directory(PUBLIC_DIR, "records.html")


@app.get("/uploads/<path:filename>")
def uploads(filename):
    return send_from_directory(UPLOAD_DIR, filename)


def _norm_currency(x: str | None) -> str | None:
    if not x:
        return None
    v = str(x).strip().upper()
    if v in {"USD", "US$", "$", "$US", "$USD"}:
        return "USD"
    if v in {"CHF", "SFR", "FR.", "FR", "CHF."}:
        return "CHF"
    if v in {"EUR", "EURO", "€"}:
        return "Euro"
    if v in {"CAD", "C$", "CA$"}:
        return "CAD"
    if v in {"USD", "CHF", "EURO", "CAD"}:
        return "Euro" if v == "EURO" else v
    return None


@app.post("/api/expenses")
def create_expense():
    try:
        if "image" not in request.files or request.files["image"].filename == "":
            return jsonify({"error": "Image is required"}), 400
        image = request.files["image"]

        # Read bytes once for hashing + saving
        img_bytes = image.read()
        if not img_bytes:
            return jsonify({"error": "Empty image"}), 400
        img_hash = hashlib.sha256(img_bytes).hexdigest()
        image.stream = BytesIO(img_bytes)

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

        currency = _norm_currency(request.form.get("currency")) or DEFAULT_CURRENCY

        name_val = (request.form.get("name") or "").strip() or f"Expense {date_str}"
        random_id = str(random.randint(1_000_000, 9_999_999))

        # Save image locally
        filename = secure_filename(image.filename)
        name_root, ext = os.path.splitext(filename)
        safe_root = name_root.replace(" ", "_") or "receipt"
        final_name = f"{int(time.time()*1000)}_{safe_root}{ext}"
        save_path = UPLOAD_DIR / final_name
        with open(save_path, "wb") as f:
            f.write(img_bytes)
        image_url = build_public_url(f"/uploads/{final_name}")
        attachment = [{"url": image_url}]

        table = get_airtable_table()
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
            "Currency": currency,
            "Hash": img_hash,
        }
        record = table.create(payload, typecast=True)
        return jsonify({"ok": True, "recordId": record.get("id"), "imageUrl": image_url})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.get("/api/expenses")
def list_expenses():
    try:
        month = request.args.get("month")
        if not month:
            today = date.today()
            month = f"{today.year:04d}-{today.month:02d}"
        formula = f"DATETIME_FORMAT({{Date added}}, 'YYYY-MM') = '{month}'"
        table = get_airtable_table()
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
            })

        # Duplicate detection by hash
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


def main():
    port = int(os.getenv("PORT", "3000"))
    app.run(host="0.0.0.0", port=port)


if __name__ == "__main__":
    main()
