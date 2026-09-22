import os
import re
import json
import time
import threading
import urllib.request
from flask import Flask, jsonify, send_from_directory, request, Response, session, redirect
from flask_cors import CORS
import openpyxl

# Load .env file automatically on local dev (Render ignores this, uses its own env panel)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # dotenv not installed — env vars must be set manually

app = Flask(__name__, static_folder="static", static_url_path="")
CORS(app)
BASE = os.path.dirname(__file__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "atmiya-foet-admin-secret-2024-xK9mP")

# ─────────────────────────────────────────────────────────────────────────────
# DEPARTMENT ADMIN CREDENTIALS
# Each entry: env-var key → { username, password, dept_filter, label }
#
# dept_filter = None  → sees ALL departments (Dean / Master view)
# dept_filter = list  → sees only rows whose "department" field matches
#
# Env vars to set in Render (per department):
#   ADMIN_MASTER_USER / ADMIN_MASTER_PASS   → Dean (sees all)
#   ADMIN_IT_USER     / ADMIN_IT_PASS       → IT HOD
#   ADMIN_CE_USER     / ADMIN_CE_PASS       → Computer Engg HOD
#   ADMIN_CIVIL_USER  / ADMIN_CIVIL_PASS    → Civil HOD
#   ADMIN_ELEC_USER   / ADMIN_ELEC_PASS     → Electrical HOD
#   ADMIN_MECH_USER   / ADMIN_MECH_PASS     → Mechanical HOD
# ─────────────────────────────────────────────────────────────────────────────
_e = os.environ.get   # shorthand

# ─────────────────────────────────────────────────────────────────────────────
# DYNAMIC CREDENTIALS & GOOGLE SHEET SYNC (Option 2)
# File: data/credentials.json
# Sheet Tab: "Users" in FOET Master Spreadsheet (Dean can manage here)
# ─────────────────────────────────────────────────────────────────────────────
_CREDENTIALS_FILE = os.path.join(BASE, "data", "credentials.json")
_cred_lock = threading.RLock()
_e = os.environ.get   # shorthand

DEPT_CRED_MAP = {
    "master": {
        "username":    _e("ADMIN_MASTER_USER") or _e("ADMIN_USERNAME") or "dean",
        "password":    _e("ADMIN_MASTER_PASS") or _e("ADMIN_PASSWORD") or "foet2024",
        "dept_filter": None,   # None = all departments (Dean)
        "label":       "FOET Master Dashboard",
        "short":       "Dean / FOET Master",
        "color":       "#002147",
    },
    "it": {
        "username":    _e("ADMIN_IT_USER") or "hod_it",
        "password":    _e("ADMIN_IT_PASS") or "it2024",
        "dept_filter": ["IT"],
        "label":       "B.Tech Information Technology - HOD Dashboard",
        "short":       "IT Department",
        "color":       "#1d4ed8",
    },
    "ce": {
        "username":    _e("ADMIN_CE_USER") or "hod_ce",
        "password":    _e("ADMIN_CE_PASS") or "ce2024",
        "dept_filter": ["Computer"],   # Contains CE and CSE programs
        "label":       "B.Tech Computer Engineering - HOD Dashboard",
        "short":       "Computer Engineering",
        "color":       "#0369a1",
    },
    "civil": {
        "username":    _e("ADMIN_CIVIL_USER") or "hod_civil",
        "password":    _e("ADMIN_CIVIL_PASS") or "civil2024",
        "dept_filter": ["Civil"],
        "label":       "B.Tech Civil Engineering - HOD Dashboard",
        "short":       "Civil Department",
        "color":       "#b45309",
    },
    "elec": {
        "username":    _e("ADMIN_ELEC_USER") or "hod_elec",
        "password":    _e("ADMIN_ELEC_PASS") or "elec2024",
        "dept_filter": ["Electrcial", "Electrical"],   # handles typo in sheet
        "label":       "B.Tech Electrical Engineering - HOD Dashboard",
        "short":       "Electrical Department",
        "color":       "#047857",
    },
    "mech": {
        "username":    _e("ADMIN_MECH_USER") or "hod_mech",
        "password":    _e("ADMIN_MECH_PASS") or "mech2024",
        "dept_filter": ["Mechanical"],
        "label":       "B.Tech Mechanical Engineering - HOD Dashboard",
        "short":       "Mechanical Department",
        "color":       "#b91c1c",
    },
}

def _load_persisted_credentials():
    """Load passwords from data/credentials.json if present."""
    if os.path.exists(_CREDENTIALS_FILE):
        try:
            with open(_CREDENTIALS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            with _cred_lock:
                for k, v in saved.items():
                    if k in DEPT_CRED_MAP and isinstance(v, dict):
                        if "username" in v and v["username"]:
                            DEPT_CRED_MAP[k]["username"] = str(v["username"]).strip()
                        if "password" in v and v["password"]:
                            DEPT_CRED_MAP[k]["password"] = str(v["password"]).strip()
        except Exception as e:
            print(f"[Auth] Could not load persisted credentials: {e}")

def _save_persisted_credentials():
    """Save current credentials to data/credentials.json."""
    try:
        os.makedirs(os.path.dirname(_CREDENTIALS_FILE), exist_ok=True)
        to_save = {}
        with _cred_lock:
            for k, v in DEPT_CRED_MAP.items():
                to_save[k] = {"username": v["username"], "password": v["password"]}
        with open(_CREDENTIALS_FILE, "w", encoding="utf-8") as f:
            json.dump(to_save, f, indent=2)
    except Exception as e:
        print(f"[Auth] Could not save persisted credentials: {e}")

_load_persisted_credentials()

def _sync_password_to_sheet(role_key, new_password):
    """Write updated password to the 'Users' tab in Google Sheets using Sheets API."""
    token = _get_service_account_token()
    if not _FOET_URL or not token:
        return
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", _FOET_URL)
    if not m:
        return
    sheet_id = m.group(1)

    try:
        import requests as req_lib
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        get_url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/Users!A:C"
        r = req_lib.get(get_url, headers=headers, timeout=10)
        if r.status_code == 200:
            rows = r.json().get("values", [])
            target_row = None
            for idx, row in enumerate(rows):
                if row and len(row) > 0 and str(row[0]).strip().lower() == role_key.lower():
                    target_row = idx + 1
                    break
            if target_row:
                put_url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/Users!C{target_row}?valueInputOption=USER_ENTERED"
                req_lib.put(put_url, headers=headers, json={"values": [[new_password]]}, timeout=10)
                print(f"[Sheet Auth] Updated password for '{role_key}' in Google Sheet row {target_row}")
                return

        append_url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/Users!A:C:append?valueInputOption=USER_ENTERED"
        u_name = DEPT_CRED_MAP.get(role_key, {}).get("username", role_key)
        req_lib.post(append_url, headers=headers, json={"values": [[role_key, u_name, new_password]]}, timeout=10)
        print(f"[Sheet Auth] Appended new password row for '{role_key}' in Google Sheet")
    except Exception as e:
        print(f"[Sheet Auth] Note: Could not update Google Sheet (ensure sheet is shared with Editor permission): {e}")


# Wrap with WhiteNoise for rock-solid production static asset serving on Render
try:
    from whitenoise import WhiteNoise
    app.wsgi_app = WhiteNoise(app.wsgi_app, root=os.path.join(os.path.dirname(__file__), "static"), prefix="")
except ImportError:
    pass

# ─────────────────────────────────────────────────────────────────────────────
# DEPARTMENT REGISTRY
# Each dept key maps to: display name, short code, color accent, data file path,
# and the ENV var name that holds the Google Sheet URL.
#
# SECURITY NOTE: Google Sheet URLs are NEVER exposed to the frontend.
# They live only in Render/Railway environment variables.
# The server fetches raw Excel data, transforms it, and returns processed JSON
# to students. Students only ever see /api/<dept>/student JSON responses.
# ─────────────────────────────────────────────────────────────────────────────
BASE = os.path.dirname(__file__)

DEPARTMENTS = {
    "it": {
        "name":      "B.Tech Information Technology",
        "short":     "IT",
        "color":     "#1d4ed8",
        "excel":     os.path.join(BASE, "data", "IT.xlsx"),
        "excel_alt": os.path.join(BASE, "data", "IT.xlsx"),
        "env_var":   "IT",
        "default_sheet_url": (
            os.environ.get("IT")
            or os.environ.get("GOOGLE_SHEET_URL_IT")
            or os.environ.get("GOOGLE_SHEET_URL")
            or ""
        ).strip(),
    },
    "civil": {
        "name":      "B.Tech Civil Engineering",
        "short":     "Civil",
        "color":     "#b45309",
        "excel":     os.path.join(BASE, "data", "Civil.xlsx"),
        "excel_alt": os.path.join(BASE, "data", "Civil.xlsx"),
        "env_var":   "Civil",
        "default_sheet_url": (
            os.environ.get("Civil")
            or os.environ.get("GOOGLE_SHEET_URL_CIVIL")
            or os.environ.get("CIVIL")
            or ""
        ).strip(),
    },
    "mech": {
        "name":      "B.Tech Mechanical Engineering",
        "short":     "Mech",
        "color":     "#b91c1c",
        "excel":     os.path.join(BASE, "data", "Mechanical.xlsx"),
        "excel_alt": os.path.join(BASE, "data", "Mechanical.xlsx"),
        "env_var":   "Mechanical",
        "default_sheet_url": (
            os.environ.get("Mechanical")
            or os.environ.get("GOOGLE_SHEET_URL_MECH")
            or os.environ.get("MECH")
            or ""
        ).strip(),
    },
    "elec": {
        "name":      "B.Tech Electrical Engineering",
        "short":     "Elec",
        "color":     "#047857",
        "excel":     os.path.join(BASE, "data", "Electrical.xlsx"),
        "excel_alt": os.path.join(BASE, "data", "Electrical.xlsx"),
        "env_var":   "Electrical",
        "default_sheet_url": (
            os.environ.get("Electrical")
            or os.environ.get("GOOGLE_SHEET_URL_ELEC")
            or os.environ.get("ELEC")
            or ""
        ).strip(),
    },
    "ce": {
        "name":      "B.Tech Computer Engineering",
        "short":     "CE",
        "color":     "#0369a1",
        "excel":     os.path.join(BASE, "data", "CE.xlsx"),
        "excel_alt": os.path.join(BASE, "data", "CE.xlsx"),
        "env_var":   "CE",
        "default_sheet_url": (
            os.environ.get("CE")
            or os.environ.get("GOOGLE_SHEET_URL_CE")
            or ""
        ).strip(),
    },
}


# ─────────────────────────────────────────────────────────────────────────────
# Per-department in-memory cache
# ─────────────────────────────────────────────────────────────────────────────
_caches = {
    dept: {
        "students":            [],
        "available_semesters": [],
        "last_modified":       0,
        "loaded_at":           0,
        "last_sync_check":     0,
        "google_sheet_url":    info["default_sheet_url"],
        "sync_status":         "Idle",
    }
    for dept, info in DEPARTMENTS.items()
}
_locks = {dept: threading.Lock() for dept in DEPARTMENTS}

# ─────────────────────────────────────────────────────────────────────────────
# FOET MASTER SHEET (admin dashboard only)
# ─────────────────────────────────────────────────────────────────────────────
_FOET_EXCEL  = os.path.join(BASE, "data", "FoET.xlsx")
_FOET_URL    = os.environ.get("FOET_MASTER_SHEET", "").strip()
_foet_cache  = {"data": None, "sync_status": "Idle", "loaded_at": 0}
_foet_lock   = threading.Lock()


def parse_foet_excel():
    """Parse FoET.xlsx and return structured data for the admin dashboard."""
    if not os.path.exists(_FOET_EXCEL):
        return None

    wb  = openpyxl.load_workbook(_FOET_EXCEL, data_only=True)
    ws  = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.active
    rows = list(ws.iter_rows(values_only=True))

    # Find header row (contains "Department")
    header_idx = None
    for i, row in enumerate(rows):
        if any(str(v or "").strip() == "Department" for v in row):
            header_idx = i
            break
    if header_idx is None:
        return None

    header = rows[header_idx]

    # Week column indices
    week_cols = [(i, str(v).strip()) for i, v in enumerate(header)
                 if v and str(v).strip().startswith("Week")]

    result_rows = []
    current_dept = None

    for row in rows[header_idx + 1:]:
        dept_val  = row[1] if len(row) > 1 else None
        prog_val  = row[2] if len(row) > 2 else None
        sem_val   = row[3] if len(row) > 3 else None
        link_val  = row[4] if len(row) > 4 else None
        count_val = row[5] if len(row) > 5 else None
        stat_val  = row[6] if len(row) > 6 else None

        # Track current department (merged cells)
        if dept_val and str(dept_val).strip() not in ("", "---"):
            current_dept = str(dept_val).strip()

        if not current_dept or not link_val or not count_val:
            continue

        # Skip summary / total rows (e.g. batch count row where sem is None or link is just a number)
        if sem_val is None or str(link_val).strip().isdigit() or "total" in str(link_val).lower():
            continue

        try:
            total = int(float(str(count_val)))
        except (ValueError, TypeError):
            continue
        if total <= 0:
            continue

        weeks = []
        for col_idx, _ in week_cols:
            val = row[col_idx] if col_idx < len(row) else None
            try:
                weeks.append(round(float(val) * 100, 2))
            except (TypeError, ValueError):
                weeks.append(None)

        # Strip trailing None and zeros (future weeks not filled yet)
        while weeks and (weeks[-1] is None or weeks[-1] == 0.0):
            weeks.pop()

        valid = [w for w in weeks if w is not None and w > 0]
        avg   = round(sum(valid) / len(valid), 2) if valid else 0

        try:
            sem = int(float(str(sem_val)))
        except (TypeError, ValueError):
            sem = 0

        result_rows.append({
            "department": current_dept,
            "program":    str(prog_val or current_dept).strip(),
            "semester":   sem,
            "link":       str(link_val).strip(),
            "total":      total,
            "status":     str(stat_val or "").strip(),
            "weeks":      weeks,
            "avg":        avg,
        })

    total_students = sum(r["total"] for r in result_rows)
    all_avgs       = [r["avg"] for r in result_rows if r["avg"] > 0]
    overall_avg    = round(sum(all_avgs) / len(all_avgs), 2) if all_avgs else 0
    below_60       = sum(1 for r in result_rows if 0 < r["avg"] < 60)
    week_labels    = [w[1] for w in week_cols]

    # Sync credentials from 'Users' or 'Admin_Users' sheet tab if Dean created one
    for sname in wb.sheetnames:
        if sname.strip().lower() in ("users", "admin_users", "user", "admin"):
            uws = wb[sname]
            updated_any = False
            with _cred_lock:
                for urow in uws.iter_rows(values_only=True):
                    if not urow or len(urow) < 3:
                        continue
                    d_key = str(urow[0] or "").strip().lower()
                    u_val = str(urow[1] or "").strip()
                    p_val = str(urow[2] or "").strip()
                    if d_key in DEPT_CRED_MAP and p_val and u_val:
                        DEPT_CRED_MAP[d_key]["username"] = u_val
                        DEPT_CRED_MAP[d_key]["password"] = p_val
                        updated_any = True
            if updated_any:
                _save_persisted_credentials()
                print(f"[Sheet Auth] Synced credentials from sheet tab: {sname}")
            break

    return {
        "rows":            result_rows,
        "week_labels":     week_labels,
        "total_students":  total_students,
        "overall_avg":     overall_avg,
        "below_60_count":  below_60,
    }


def _sync_foet_bg():
    """Background thread: download master sheet and reload cache."""
    token = _get_service_account_token()
    if not _FOET_URL or not token:
        with _foet_lock:
            _foet_cache["sync_status"] = "No URL or service account configured"
        return

    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", _FOET_URL)
    sheet_id = m.group(1) if m else ""

    try:
        import requests as req_lib
        headers = {"Authorization": f"Bearer {token}", "Accept": "*/*"}
        content = None

        if sheet_id:
            url = (f"https://www.googleapis.com/drive/v3/files/{sheet_id}/export"
                   f"?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            resp = req_lib.get(url, headers=headers, timeout=30)
            if resp.status_code == 200 and len(resp.content) > 1000:
                content = resp.content

        if not content:
            export_url = get_export_url(_FOET_URL)
            resp = req_lib.get(export_url, headers=headers, timeout=30)
            if resp.status_code == 200 and len(resp.content) > 1000:
                content = resp.content

        if content:
            os.makedirs(os.path.dirname(_FOET_EXCEL), exist_ok=True)
            with open(_FOET_EXCEL, "wb") as f:
                f.write(content)
            data = parse_foet_excel()
            with _foet_lock:
                _foet_cache["data"]        = data
                _foet_cache["loaded_at"]   = time.time()
                _foet_cache["sync_status"] = f"Synced at {time.strftime('%H:%M:%S')}"
        else:
            with _foet_lock:
                _foet_cache["sync_status"] = "Sync failed: could not download"
    except Exception as e:
        with _foet_lock:
            _foet_cache["sync_status"] = f"Sync error: {e}"



# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def get_export_url(raw_url):
    """Convert any Google Sheets shareable link into a direct .xlsx download URL."""
    if not raw_url:
        return ""
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", raw_url)
    if m:
        sheet_id = m.group(1)
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    if "export?format=xlsx" in raw_url:
        return raw_url
    return raw_url


def _get_service_account_token():
    """
    Load service account credentials from GOOGLE_SERVICE_ACCOUNT_JSON env var
    or a local credentials JSON file, and return a Bearer token.
    Returns None if not configured or invalid.
    """
    import glob
    sa_info = None
    sa_json_str = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()

    if sa_json_str:
        try:
            sa_info = json.loads(sa_json_str)
        except Exception as e:
            print(f"[Service Account] Failed to parse GOOGLE_SERVICE_ACCOUNT_JSON: {e}")
    else:
        creds_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "").strip()
        if not creds_path:
            candidates = glob.glob("atmiya-*.json") + glob.glob("service-account*.json")
            for candidate in candidates:
                if os.path.isfile(candidate):
                    creds_path = candidate
                    break
        if creds_path and os.path.exists(creds_path):
            try:
                with open(creds_path, "r", encoding="utf-8") as f:
                    sa_info = json.load(f)
            except Exception as e:
                print(f"[Service Account] Failed to read {creds_path}: {e}")

    if not sa_info:
        return None

    try:
        import google.oauth2.service_account as sa_module
        import google.auth.transport.requests as ga_requests
        creds = sa_module.Credentials.from_service_account_info(
            sa_info,
            scopes=[
                "https://www.googleapis.com/auth/drive",
                "https://www.googleapis.com/auth/spreadsheets",
            ]
        )
        creds.refresh(ga_requests.Request())
        return creds.token
    except Exception as e:
        print(f"[Service Account] Failed to get token: {e}")
        return None


def sync_google_sheet(dept_key, sheet_url=None):
    """
    Download live .xlsx from Google Sheets and save to data/<DEPT>.xlsx.

    SECURITY:
    - If service account credentials exist (file or env var) → authenticates as the service account.
      The sheet can stay RESTRICTED (shared with foet-portal@atmiya-foet-portal.iam.gserviceaccount.com).
    - Otherwise falls back to plain URL download (sheet must be 'Anyone with link can view').
    - Sheet URL is NEVER sent to the browser.
    """
    cache     = _caches[dept_key]
    dept_info = DEPARTMENTS[dept_key]
    token     = _get_service_account_token()

    target_url = sheet_url or cache.get("google_sheet_url") or dept_info["default_sheet_url"]

    # Auto-discovery fallback: If no URL is explicitly configured, check Drive for files shared with SA
    if not target_url and token:
        try:
            import requests as req_lib
            r = req_lib.get(
                "https://www.googleapis.com/drive/v3/files?q=trashed=false&fields=files(id,name,mimeType)",
                headers={"Authorization": f"Bearer {token}"},
                timeout=10
            )
            if r.status_code == 200:
                files = r.json().get("files", [])
                for f in files:
                    fname = f.get("name", "").lower()
                    if dept_key in fname or dept_info["short"].lower() in fname:
                        target_url = f"https://docs.google.com/spreadsheets/d/{f['id']}/edit"
                        cache["google_sheet_url"] = target_url
                        print(f"[Sync] Auto-discovered Google Drive file '{f.get('name')}' ({f['id']}) for {dept_key}")
                        break
        except Exception as e:
            print(f"[Sync] Auto-discovery error: {e}")

    if not target_url:
        cache["sync_status"] = "No Google Sheet URL configured"
        return False, "No Google Sheet URL configured for this department."

    export_url = get_export_url(target_url)
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", target_url)
    sheet_id = m.group(1) if m else ""

    try:
        import requests as req_lib
        content = None
        status_code = None

        if token:
            headers = {
                "Authorization": f"Bearer {token}",
                "Accept": "*/*",
            }

            # 1. Drive API export — correct method for service accounts on native Google Sheets
            if sheet_id:
                try:
                    url_export = (
                        f"https://www.googleapis.com/drive/v3/files/{sheet_id}/export"
                        f"?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    resp = req_lib.get(url_export, headers=headers, timeout=30)
                    if resp.status_code == 200 and len(resp.content) > 1000:
                        content = resp.content
                        status_code = 200
                    else:
                        status_code = resp.status_code
                except Exception:
                    pass

            # 2. docs.google.com export URL (fallback)
            if (not content or len(content) <= 1000) and export_url:
                try:
                    resp = req_lib.get(export_url, headers=headers, timeout=30)
                    if resp.status_code == 200 and len(resp.content) > 1000:
                        content = resp.content
                        status_code = 200
                    else:
                        status_code = resp.status_code
                except Exception:
                    pass

            # 3. Drive API alt=media (for uploaded xlsx files, not native Sheets)
            if (not content or len(content) <= 1000) and sheet_id:
                try:
                    url_media = f"https://www.googleapis.com/drive/v3/files/{sheet_id}?alt=media"
                    resp = req_lib.get(url_media, headers=headers, timeout=30)
                    if resp.status_code == 200 and len(resp.content) > 1000:
                        content = resp.content
                        status_code = 200
                    elif not status_code:
                        status_code = resp.status_code
                except Exception:
                    pass

        else:
            # ── Fallback: plain URL download (requires public sheet) ────────
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                "Accept": "*/*",
            }
            req_obj = urllib.request.Request(export_url, headers=headers)
            with urllib.request.urlopen(req_obj, timeout=15) as response:
                content = response.read()
            status_code = 200

        if status_code in (403, 404):
            err_msg = (
                f"Google Drive returned HTTP {status_code}. "
                f"Make sure the sheet/file is shared with: foet-portal@atmiya-foet-portal.iam.gserviceaccount.com"
            )
            cache["sync_status"] = f"Failed: HTTP {status_code} (Share with service account)"
            return False, err_msg

        if status_code != 200 or not content:
            err_msg = f"Could not download Google Sheet: HTTP {status_code}"
            cache["sync_status"] = f"Failed: HTTP {status_code}"
            return False, err_msg

        if len(content) > 1000:
            excel_path = dept_info["excel"]
            os.makedirs(os.path.dirname(excel_path), exist_ok=True)
            with open(excel_path, "wb") as f:
                f.write(content)
            now_str = time.strftime("%H:%M:%S")
            cache["sync_status"] = f"Synced live ({len(content)} bytes at {now_str})"
            return True, f"Live Google Sheet synced successfully ({len(content)} bytes)!"
        else:
            cache["sync_status"] = "Failed: Downloaded file is too small or invalid"
            return False, "Downloaded file is too small or invalid."

    except Exception as e:
        cache["sync_status"] = f"Sync failed: {str(e)}"
        return False, f"Could not download Google Sheet: {str(e)}"


def extract_semester_from_title(title):
    """Pull semester number out of the sheet title row."""
    if not title:
        return None
    m = re.search(r"SEMESTER\s*[-–]\s*(\d+)", str(title), re.IGNORECASE)
    if m:
        return int(m.group(1))
    return None


def parse_excel(dept_key):
    """
    Parse the Excel file for a department row by row.
    A single sheet can contain multiple semester sections — each section starts
    with a title row like "B.TECH. INFORMATION TECHNOLOGY SEMESTER - 5 ATTENDANCE REPORT".
    """
    dept_info = DEPARTMENTS[dept_key]
    excel_path = dept_info["excel"] if os.path.exists(dept_info["excel"]) else dept_info["excel_alt"]

    if not os.path.exists(excel_path):
        return [], []

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    all_students = []
    semesters_found = set()

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        current_semester = None
        in_data = False

        for row in rows:
            first_cell = str(row[0]).strip() if row[0] is not None else ""

            sem = extract_semester_from_title(first_cell)
            if sem:
                current_semester = sem
                semesters_found.add(sem)
                in_data = False
                continue

            if row[0] == "Div" and row[1] == "Batch":
                in_data = True
                continue

            if current_semester is None or not in_data:
                continue

            div = row[0]
            name = row[5]
            if not name or div in (None, "Div", ""):
                continue

            try:
                roll_no = int(row[2]) if row[2] is not None else None
            except (ValueError, TypeError):
                continue
            if not isinstance(name, str):
                continue

            try:
                reg_no = str(int(row[3])) if row[3] else ""
                enr_no = str(int(row[4])) if row[4] else ""
            except (ValueError, TypeError):
                continue

            status        = row[6]  if row[6]  is not None else "N/A"
            points        = row[7]  if row[7]  is not None else 0
            total_hours   = row[8]  if row[8]  is not None else 0
            comp_required = row[9]  if row[9]  is not None else 0
            comp_completed= row[10] if row[10] is not None else 0

            weeks = []
            # Dynamic: read ALL columns from 11 onwards as week data.
            # Never break on an empty cell — gaps in the sheet are valid.
            # After reading, strip trailing NOT_AVAILABLE entries so future
            # empty weeks don't show as N/A.
            week_num = 1
            for w in range(11, len(row)):
                val = row[w]
                if val is None:
                    weeks.append({"week": week_num, "attendance": None, "status": "NOT_AVAILABLE"})
                else:
                    try:
                        pct = round(float(val) * 100, 2)
                    except (ValueError, TypeError):
                        weeks.append({"week": week_num, "attendance": None, "status": "NOT_AVAILABLE"})
                        week_num += 1
                        continue
                    if pct >= 80:
                        wk_status = "OK"
                    elif pct >= 60:
                        wk_status = "WARNING"
                    else:
                        wk_status = "PENDING"
                    weeks.append({"week": week_num, "attendance": pct, "status": wk_status})
                week_num += 1

            # Strip trailing NOT_AVAILABLE weeks (future weeks not yet updated)
            while weeks and weeks[-1]["status"] == "NOT_AVAILABLE":
                weeks.pop()

            attended = [w for w in weeks if w["attendance"] is not None]
            avg_att  = round(sum(w["attendance"] for w in attended) / len(attended), 2) if attended else 0.0
            total_req = 24 * len(attended) if attended else 0

            all_students.append({
                "department":            dept_info["short"],
                "semester":              current_semester,
                "div":                   str(div).strip(),
                "batch":                 str(row[1]).strip() if row[1] else "",
                "roll_no":               roll_no,
                "reg_no":                reg_no,
                "enr_no":                enr_no,
                "name":                  name.strip(),
                "status":                str(status).strip(),
                "points":                int(points) if points else 0,
                "total_hours_attended":  int(total_hours) if total_hours else 0,
                "total_hours_required":  total_req,
                "comp_required":         int(comp_required) if comp_required else 0,
                "comp_completed":        int(comp_completed) if comp_completed else 0,
                "attendance_pct":        avg_att,
                "weeks":                 weeks,
            })

    return all_students, sorted(semesters_found)


def get_students(dept_key):
    """Return cached students for a department, refreshing if Excel changed."""
    now   = time.time()
    cache = _caches[dept_key]
    info  = DEPARTMENTS[dept_key]

    # Auto-sync from Google Sheets every 60 s if URL is configured
    if cache.get("google_sheet_url") and (now - cache["last_sync_check"] > 60):
        cache["last_sync_check"] = now
        threading.Thread(target=sync_google_sheet, args=(dept_key,), daemon=True).start()

    excel_path = info["excel"] if os.path.exists(info["excel"]) else info["excel_alt"]
    try:
        mtime = os.path.getmtime(excel_path)
    except OSError:
        mtime = 0

    with _locks[dept_key]:
        if mtime != cache["last_modified"] or not cache["students"]:
            students, semesters = parse_excel(dept_key)
            cache["students"]            = students
            cache["available_semesters"] = semesters
            cache["last_modified"]       = mtime
            cache["loaded_at"]           = time.time()

    return cache["students"], cache["available_semesters"]


def has_data(dept_key):
    """Return True if a data file exists for this department."""
    info = DEPARTMENTS[dept_key]
    return os.path.exists(info["excel"]) or os.path.exists(info["excel_alt"])


# ─────────────────────────────────────────────────────────────────────────────
# STATIC PAGE ROUTES
# ─────────────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory("static", "index.html")


# One route per department — all serve the same dept.html template
DEPT_ROUTES = list(DEPARTMENTS.keys())

@app.route("/<dept_slug>")
def dept_page(dept_slug):
    if dept_slug in DEPARTMENTS:
        return send_from_directory("static", "dept.html")
    return "Page not found", 404


@app.route("/style.css")
def css():
    return send_from_directory("static", "style.css")

@app.route("/app.js")
def js():
    return send_from_directory("static", "app.js")

@app.route("/master.js")
def master_js():
    return send_from_directory("static", "master.js")

@app.route("/atmiyalogonaac.png")
def naac_logo():
    return send_from_directory("static", "atmiyalogonaac.png")

@app.route("/favicon.ico")
def favicon():
    return Response(status=204)


# ─────────────────────────────────────────────────────────────────────────────
# API: Master department list
# ─────────────────────────────────────────────────────────────────────────────
@app.route("/api/departments")
def api_departments():
    result = []
    for key, info in DEPARTMENTS.items():
        data_available = has_data(key)
        cache = _caches[key]
        has_sync = bool(cache.get("google_sheet_url") or info.get("default_sheet_url"))
        result.append({
            "key":             key,
            "name":            info["name"],
            "short":           info["short"],
            "color":           info["color"],
            "url":             f"/{key}",
            "data_available":  data_available,
            "sync_configured": has_sync,
        })
    return jsonify({"departments": result})


# ─────────────────────────────────────────────────────────────────────────────
# API: Per-department endpoints
# ─────────────────────────────────────────────────────────────────────────────
@app.route("/api/<dept>/semesters")
def api_semesters(dept):
    dept = dept.lower()
    if dept not in DEPARTMENTS:
        return jsonify({"error": f"Unknown department: {dept}"}), 404
    if not has_data(dept):
        return jsonify({"department": DEPARTMENTS[dept]["short"], "semesters": [], "data_available": False})
    _, available = get_students(dept)
    return jsonify({"department": DEPARTMENTS[dept]["short"], "semesters": available, "data_available": True})


@app.route("/api/<dept>/student")
def api_student(dept):
    dept = dept.lower()
    if dept not in DEPARTMENTS:
        return jsonify({"error": f"Unknown department: {dept}"}), 404

    if not has_data(dept):
        return jsonify({
            "error": f"Student record not found. Please verify your Registration / Enrollment number and semester, or contact your department coordinator."
        }), 404

    query          = request.args.get("q", "").strip()
    semester_param = request.args.get("semester", "").strip()

    if not query:
        return jsonify({"error": "Please enter your Registration or Enrollment number."}), 400
    if not semester_param:
        return jsonify({"error": "Please select your semester."}), 400

    try:
        semester_num = int(semester_param)
    except ValueError:
        return jsonify({"error": "Invalid semester selected."}), 400

    students, available_semesters = get_students(dept)
    dept_name = DEPARTMENTS[dept]["name"]

    if semester_num not in available_semesters:
        return jsonify({
            "error": f"Semester {semester_num} data is not available for {dept_name}. "
                     f"Available: {', '.join(str(s) for s in available_semesters)}."
        }), 404

    found = None
    for s in students:
        if s["semester"] != semester_num:
            continue
        if s["reg_no"] == query or s["enr_no"] == query:
            found = s
            break

    if not found:
        other = next(
            (s for s in students if (s["reg_no"] == query or s["enr_no"] == query)
             and s["semester"] != semester_num),
            None
        )
        if other:
            return jsonify({
                "error": f"This number belongs to Semester {other['semester']}, "
                         f"not Semester {semester_num}. Please select the correct semester."
            }), 404
        return jsonify({"error": f"No student found with this number in {dept_name}."}), 404

    return jsonify(found)


@app.route("/api/<dept>/sync", methods=["GET", "POST"])
def api_sync(dept):
    dept = dept.lower()
    if dept not in DEPARTMENTS:
        return jsonify({"error": f"Unknown department: {dept}"}), 404

    cache = _caches[dept]
    if request.method == "POST":
        body = request.get_json(silent=True) or {}
        new_url = body.get("url") or request.args.get("url", "").strip()
        if new_url:
            cache["google_sheet_url"] = new_url

    url = cache.get("google_sheet_url")
    if not url:
        return jsonify({
            "status":         "not_configured",
            "message":        f"No Google Sheet configured for {dept.upper()}. Serving local data.",
            "total_students": len(cache.get("students", [])),
            "semesters":      cache.get("available_semesters", []),
        }), 200

    # Run in background — avoid Render's 30 s request timeout
    def _bg_sync():
        ok, msg = sync_google_sheet(dept, url)
        if ok:
            with _locks[dept]:
                students, semesters = parse_excel(dept)
                cache["students"]            = students
                cache["available_semesters"] = semesters
                cache["loaded_at"]           = time.time()
            cache["sync_status"] = f"Synced OK at {time.strftime('%H:%M:%S')}"
        else:
            cache["sync_status"] = f"Sync failed: {msg}"

    threading.Thread(target=_bg_sync, daemon=True).start()

    return jsonify({
        "status":         "started",
        "message":        "Sync running in background — fresh data ready in ~10 s.",
        "total_students": len(cache.get("students", [])),
        "semesters":      cache.get("available_semesters", []),
    })



@app.route("/api/<dept>/status")
def api_status(dept):
    dept = dept.lower()
    if dept not in DEPARTMENTS:
        return jsonify({"error": f"Unknown department: {dept}"}), 404
    students, semesters = get_students(dept) if has_data(dept) else ([], [])
    cache = _caches[dept]
    return jsonify({
        "department":            DEPARTMENTS[dept]["short"],
        "name":                  DEPARTMENTS[dept]["name"],
        "data_available":        has_data(dept),
        "total_students":        len(students),
        "available_semesters":   semesters,
        "loaded_at":             cache["loaded_at"],
        "google_sheet_url":      cache.get("google_sheet_url", ""),
        "sync_status":           cache.get("sync_status", ""),
        "service_account_email": "foet-portal@atmiya-foet-portal.iam.gserviceaccount.com",
    })


# ─────────────────────────────────────────────────────────────────────────────
# Backward-compatible aliases (keep old IT-only URLs working)
# ─────────────────────────────────────────────────────────────────────────────
@app.route("/api/semesters")
def compat_semesters():
    return api_semesters("it")

@app.route("/api/student")
def compat_student():
    return api_student("it")

@app.route("/api/sync", methods=["GET", "POST"])
def compat_sync():
    return api_sync("it")

@app.route("/api/status")
def compat_status():
    return api_status("it")



# ─────────────────────────────────────────────────────────────────────────────
# Cache-busting: rewrite ?v= in every HTML response so browsers always
# load the latest JS/CSS without needing a manual hard-refresh.
# ─────────────────────────────────────────────────────────────────────────────
_ASSET_VERSION = str(int(time.time()))

@app.after_request
def bust_cache(response):
    if "text/html" in response.content_type:
        try:
            body = response.get_data(as_text=True)
            body = re.sub(r'\?v=\d+', f'?v={_ASSET_VERSION}', body)
            response.set_data(body)
        except RuntimeError:
            pass  # WhiteNoise passthrough mode — skip
    return response



# ─────────────────────────────────────────────────────────────────────────────
# STARTUP SYNC
# On server boot, immediately sync all departments that have a sheet URL
# configured. Runs in background threads so startup isn't blocked.
# ─────────────────────────────────────────────────────────────────────────────
def _startup_sync():
    """Wait a moment for Flask to fully start, then sync all configured depts."""
    time.sleep(5)
    print("[Startup] Beginning initial sync for all configured departments...")
    for dept_key, info in DEPARTMENTS.items():
        url = _caches[dept_key].get("google_sheet_url") or info.get("default_sheet_url", "")
        if url:
            print(f"[Startup] Syncing {info['short']} ({dept_key})...")
            ok, msg = sync_google_sheet(dept_key)
            print(f"[Startup] {info['short']}: {'✓' if ok else '✗'} {msg}")
        else:
            print(f"[Startup] {info['short']}: Skipped (no sheet URL configured)")


def _load_foet_on_boot():
    time.sleep(6)
    data = parse_foet_excel()
    if data:
        with _foet_lock:
            _foet_cache["data"]      = data
            _foet_cache["loaded_at"] = time.time()
            _foet_cache["sync_status"] = "Loaded from file"
    if _FOET_URL:
        threading.Thread(target=_sync_foet_bg, daemon=True).start()


if os.environ.get("TESTING") != "1":
    _startup_thread = threading.Thread(target=_startup_sync, daemon=True)
    _startup_thread.start()
    threading.Thread(target=_load_foet_on_boot, daemon=True).start()


# ─────────────────────────────────────────────────────────────────────────────
# ADMIN ROUTES  (Faculty / HOD / Dean only — session-protected)
# ─────────────────────────────────────────────────────────────────────────────
# ADMIN ROUTES  (Faculty / HOD / Dean only — session + dept-filtered)
# ─────────────────────────────────────────────────────────────────────────────
ADMIN_PAGES = os.path.join(BASE, "admin_pages")

def _admin_authed():
    return session.get("admin_logged_in") is True

def _admin_role():
    """Return the DEPT_CRED_MAP entry for the current session user."""
    role_key = session.get("admin_role", "master")
    return DEPT_CRED_MAP.get(role_key, DEPT_CRED_MAP["master"])

def _filter_foet_data(data, dept_filter):
    """Return a copy of data with rows filtered to dept_filter (or all if None)."""
    if not data:
        return data
    if dept_filter is None:
        return data   # master sees everything
    filtered = [r for r in data["rows"] if r["department"] in dept_filter]
    if not filtered:
        return {**data, "rows": [], "total_students": 0, "overall_avg": 0, "below_60_count": 0}
    total   = sum(r["total"] for r in filtered)
    avgs    = [r["avg"] for r in filtered if r["avg"] > 0]
    overall = round(sum(avgs) / len(avgs), 2) if avgs else 0
    below60 = sum(1 for r in filtered if 0 < r["avg"] < 60)
    return {**data, "rows": filtered, "total_students": total,
            "overall_avg": overall, "below_60_count": below60}


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    dept_param = request.args.get("dept", "").strip().lower()
    if dept_param in DEPT_CRED_MAP and dept_param != "master":
        return redirect(f"/admin/{dept_param}")

    if _admin_authed():
        return redirect("/admin")
    error = ""
    if request.method == "POST":
        u = request.form.get("username", "").strip()
        p = request.form.get("password", "").strip()
        matched = None
        for role_key, cred in DEPT_CRED_MAP.items():
            if cred["username"] and u == cred["username"] and p == cred["password"]:
                matched = role_key
                break
        if matched:
            session["admin_logged_in"] = True
            session["admin_role"]      = matched
            session.permanent          = False
            return redirect("/admin")
        error = "Invalid username or password."
    with open(os.path.join(ADMIN_PAGES, "login.html"), encoding="utf-8") as f:
        html = f.read()
    err_html = f'<div class="error-msg">{error}</div>' if error else ""
    html = html.replace("{{ERROR_BLOCK}}", err_html)
    html = html.replace("{{PORTAL_TITLE}}", "FOET Faculty &amp; HOD Portal")
    html = html.replace("{{PORTAL_SUB}}", "Sign in with your department credentials to view your semester dashboard")
    html = html.replace("{{ACTION_URL}}", "/admin/login")
    html = html.replace("{{ACTIVE_DEPT}}", "master")
    return Response(html, content_type="text/html")


@app.route("/admin/<dept_slug>", methods=["GET", "POST"])
def admin_dept_login(dept_slug):
    dept_slug = dept_slug.lower()
    if dept_slug not in DEPT_CRED_MAP:
        return redirect("/admin/login")

    # If already logged into this department (or master), open dashboard directly
    if _admin_authed():
        curr = session.get("admin_role")
        if curr == dept_slug or curr == "master":
            return redirect("/admin")

    error = ""
    cred = DEPT_CRED_MAP[dept_slug]
    if request.method == "POST":
        u = request.form.get("username", "").strip()
        p = request.form.get("password", "").strip()
        # Verify credentials for this specific department
        if cred["username"] and u == cred["username"] and p == cred["password"]:
            session["admin_logged_in"] = True
            session["admin_role"]      = dept_slug
            session.permanent          = False
            return redirect("/admin")
        # Also allow dean to login from here into this department's view
        master_cred = DEPT_CRED_MAP["master"]
        if u == master_cred["username"] and p == master_cred["password"]:
            session["admin_logged_in"] = True
            session["admin_role"]      = dept_slug
            session.permanent          = False
            return redirect("/admin")
        error = f"Invalid username or password for {cred['short']}."

    with open(os.path.join(ADMIN_PAGES, "login.html"), encoding="utf-8") as f:
        html = f.read()
    err_html = f'<div class="error-msg">{error}</div>' if error else ""
    html = html.replace("{{ERROR_BLOCK}}", err_html)
    html = html.replace("{{PORTAL_TITLE}}", f"{cred['short']} — Faculty Login")
    html = html.replace("{{PORTAL_SUB}}", f"Sign in to view {cred['label']}")
    html = html.replace("{{ACTION_URL}}", f"/admin/{dept_slug}")
    html = html.replace("{{ACTIVE_DEPT}}", dept_slug)
    return Response(html, content_type="text/html")


@app.route("/<dept_slug>/admin")
def dept_admin_shortcut(dept_slug):
    dept_slug = dept_slug.lower()
    if dept_slug in DEPT_CRED_MAP:
        return redirect(f"/admin/{dept_slug}")
    return redirect("/admin/login")


@app.route("/admin/logout")
def admin_logout():
    session.clear()
    return redirect("/admin/login")


@app.route("/admin")
def admin_dashboard():
    if not _admin_authed():
        return redirect("/admin/login")
    with open(os.path.join(ADMIN_PAGES, "dashboard.html"), encoding="utf-8") as f:
        html = f.read()
    return Response(html, content_type="text/html")


@app.route("/api/admin/session")
def api_admin_session():
    """Return current session role info for the dashboard to show correct title."""
    if not _admin_authed():
        return jsonify({"error": "Unauthorized"}), 401
    role_key = session.get("admin_role", "master")
    role = DEPT_CRED_MAP.get(role_key, DEPT_CRED_MAP["master"])
    return jsonify({
        "role_key":  role_key,
        "label":     role["label"],
        "short":     role["short"],
        "color":     role.get("color", "#002147"),
        "is_master": role_key == "master",
    })


@app.route("/api/admin/data")
def api_admin_data():
    if not _admin_authed():
        return jsonify({"error": "Unauthorized"}), 401
    data = _foet_cache.get("data")
    if not data:
        data = parse_foet_excel()
        if data:
            with _foet_lock:
                _foet_cache["data"]      = data
                _foet_cache["loaded_at"] = time.time()
    role        = _admin_role()
    dept_filter = role["dept_filter"]
    filtered    = _filter_foet_data(data, dept_filter)
    return jsonify({
        "data":        filtered,
        "sync_status": _foet_cache.get("sync_status", "Idle"),
        "loaded_at":   _foet_cache.get("loaded_at", 0),
    })


@app.route("/api/admin/sync", methods=["POST"])
def api_admin_sync():
    if not _admin_authed():
        return jsonify({"error": "Unauthorized"}), 401
    with _foet_lock:
        _foet_cache["sync_status"] = "Syncing..."
    threading.Thread(target=_sync_foet_bg, daemon=True).start()
    return jsonify({"status": "started", "message": "Sync running in background — fresh data ready in ~10 s."})


@app.route("/api/admin/change-password", methods=["POST"])
def api_admin_change_password():
    if not _admin_authed():
        return jsonify({"success": False, "error": "Unauthorized"}), 401

    data = request.get_json(force=True, silent=True) or {}
    curr_pass = str(data.get("current_password") or "").strip()
    new_pass  = str(data.get("new_password") or "").strip()
    conf_pass = str(data.get("confirm_password") or "").strip()

    role_key = session.get("admin_role", "master")
    cred = DEPT_CRED_MAP.get(role_key)
    if not cred:
        return jsonify({"success": False, "error": "Invalid session role"}), 400

    if not curr_pass or curr_pass != cred["password"]:
        return jsonify({"success": False, "error": "Current password does not match."}), 400

    if not new_pass or len(new_pass) < 6:
        return jsonify({"success": False, "error": "New password must be at least 6 characters long."}), 400

    if new_pass != conf_pass:
        return jsonify({"success": False, "error": "New password and confirmation do not match."}), 400

    with _cred_lock:
        DEPT_CRED_MAP[role_key]["password"] = new_pass
        _save_persisted_credentials()

    # Attempt to write back to Google Sheet Users tab in background
    threading.Thread(target=_sync_password_to_sheet, args=(role_key, new_pass), daemon=True).start()

    return jsonify({
        "success": True,
        "message": f"Password for {cred['short']} updated successfully."
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"Starting Atmiya University — FOET Attendance Portal on port {port}")
    app.run(host="0.0.0.0", port=port, debug=False)
