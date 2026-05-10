import json
import os
import threading
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional

import jwt
from fastapi import FastAPI, Header, HTTPException
from pydantic import BaseModel

app = FastAPI(title="ArchLens Credits API")

ARCHLENS_WEBHOOK_SECRET = os.getenv(
    "ARCHLENS_WEBHOOK_SECRET",
    "archlens_secure_2026_SYDS_92838"
)
ARCHLENS_SHARED_SECRET = os.getenv(
    "ARCHLENS_SHARED_SECRET",
    "ArchLens-SYDS-2026-very-long-random-private-secret-839201",
)

# Persistent local storage.
# Important: For full production, connect Render PostgreSQL/Supabase.
# This JSON store fixes the in-memory reset issue during normal app reloads and worker restarts
# where the file system remains available.
DATA_DIR = Path(os.getenv("ARCHLENS_DATA_DIR", "data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)
CREDITS_FILE = DATA_DIR / "credits_store.json"

STORE_LOCK = threading.Lock()


class CreditRequest(BaseModel):
    email: str
    credits: int
    orderId: str
    source: str = "wix_store"


class DeductCreditRequest(BaseModel):
    email: str
    credits: int
    reportId: str = ""
    exportType: str = ""
    source: str = "archlens_download"


class AdminCreditRequest(BaseModel):
    email: str
    credits: int
    reason: str = "manual_adjustment"


class AdminAdjustCreditRequest(BaseModel):
    email: str
    action: str
    credits: int
    reason: str
    admin_email: Optional[str] = ""


class ReportRecordRequest(BaseModel):
    email: str
    plan: str = ""
    report_id: str = ""
    project_name: str = ""
    project_address: str = ""
    report_type: str = ""
    credits_used: int = 0
    download_path: str = ""
    status: str = "generated"


class UserActivityRequest(BaseModel):
    email: str
    plan: str = ""
    status: str = "active"


class AdminUserStatusRequest(BaseModel):
    email: str
    status: str
    reason: str


def normalise_email(email: str) -> str:
    return str(email or "").strip().lower()


def find_email_in_payload(payload) -> str:
    if isinstance(payload, dict):
        for key in ("email", "primaryEmail", "primary_email", "loginEmail", "memberEmail", "contactEmail"):
            value = payload.get(key)
            if is_valid_email(value):
                return str(value).strip()
        for value in payload.values():
            found = find_email_in_payload(value)
            if found:
                return found
    elif isinstance(payload, list):
        for value in payload:
            found = find_email_in_payload(value)
            if found:
                return found
    elif isinstance(payload, str):
        value = payload.strip()
        if is_valid_email(value):
            return value
    return ""


def is_valid_email(email: str) -> bool:
    clean = normalise_email(email)
    return "@" in clean and "." in clean.split("@")[-1]


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def configured_admin_emails() -> List[str]:
    raw = os.getenv("ADMIN_EMAILS", "")
    return [
        email.strip().lower()
        for email in raw.split(",")
        if email.strip()
    ]


@app.on_event("startup")
def log_admin_config() -> None:
    print(f"ArchLens Admin: {len(configured_admin_emails())} admin email(s) loaded from ADMIN_EMAILS.")


def default_store() -> Dict:
    return {
        "users": {},
        "processed_orders": [],
        "processed_unlocks": [],
        "transactions": [],
        "reports": [],
        "user_meta": {},
        "audit_log": [],
        "errors": [],
    }


def load_store() -> Dict:
    if not CREDITS_FILE.exists():
        return default_store()

    try:
        with CREDITS_FILE.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return default_store()

    if not isinstance(data, dict):
        return default_store()

    data.setdefault("users", {})
    data.setdefault("processed_orders", [])
    data.setdefault("processed_unlocks", [])
    data.setdefault("transactions", [])
    data.setdefault("reports", [])
    data.setdefault("user_meta", {})
    data.setdefault("audit_log", [])
    data.setdefault("errors", [])
    return data


def save_store(data: Dict) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    temp_file = CREDITS_FILE.with_suffix(".tmp")
    with temp_file.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, sort_keys=True)
    temp_file.replace(CREDITS_FILE)


def get_balance_from_store(email: str) -> int:
    clean_email = normalise_email(email)
    with STORE_LOCK:
        store = load_store()
        return int(store["users"].get(clean_email, 0) or 0)


def add_transaction(store: Dict, email: str, amount: int, balance_after: int, reason: str, source: str, reference: str = "") -> None:
    transactions: List[Dict] = store.setdefault("transactions", [])
    transactions.insert(
        0,
        {
            "email": email,
            "amount": int(amount),
            "balance_after": int(balance_after),
            "reason": reason,
            "source": source,
            "reference": reference,
            "timestamp": utc_now(),
        },
    )
    store["transactions"] = transactions[:500]


def update_user_meta(store: Dict, email: str, plan: str = "", status: str = "") -> None:
    clean_email = normalise_email(email)
    if not clean_email:
        return
    meta = store.setdefault("user_meta", {}).get(clean_email, {})
    if plan:
        meta["plan"] = plan
    if status:
        meta["status"] = status
    else:
        meta.setdefault("status", "active")
    meta["last_activity"] = utc_now()
    store.setdefault("user_meta", {})[clean_email] = meta


def email_from_admin_token(authorization: str = "", x_archlens_token: str = "", query_token: str = "") -> str:
    header_token = str(x_archlens_token or "").strip()
    query_token_value = str(query_token or "").strip()
    auth_value = str(authorization or "").strip()
    token_value = header_token
    if not token_value and auth_value.lower().startswith("bearer "):
        token_value = auth_value.split(" ", 1)[1].strip()
    if not token_value:
        token_value = query_token_value
    print(
        "ArchLens Admin diagnostics request:",
        {
            "authorization_header_present": bool(auth_value),
            "x_archlens_token_present": bool(header_token),
            "query_token_present": bool(query_token_value),
            "token_received": bool(token_value),
        },
    )
    if not token_value:
        raise HTTPException(status_code=401, detail="Missing admin auth token")
    try:
        payload = jwt.decode(token_value, ARCHLENS_SHARED_SECRET, algorithms=["HS256"])
    except Exception as exc:
        print("ArchLens Admin token decode failed:", str(exc))
        raise HTTPException(status_code=401, detail="Invalid admin bearer token")
    clean_email = normalise_email(find_email_in_payload(payload))
    print("ArchLens Admin authenticated email extracted:", clean_email or "not found")
    if not is_valid_email(clean_email):
        raise HTTPException(status_code=401, detail="Admin bearer token does not contain a valid email")
    return clean_email


def verify_admin_access(x_archlens_secret: str, authorization: str = "", x_archlens_token: str = "", token: str = "") -> str:
    verify_secret(x_archlens_secret)
    clean_admin = email_from_admin_token(authorization, x_archlens_token, token)
    admins = configured_admin_emails()
    is_allowed = bool(clean_admin and clean_admin in admins)
    print(
        "ArchLens Admin validation:",
        {
            "authenticated_email": clean_admin,
            "admin_emails": admins,
            "is_admin": is_allowed,
        },
    )
    if not is_allowed:
        raise HTTPException(status_code=403, detail="Authenticated user email is not listed in ADMIN_EMAILS")
    return clean_admin


def admin_session_payload(authorization: str = "", x_archlens_token: str = "", token: str = "") -> Dict:
    clean_admin = email_from_admin_token(authorization, x_archlens_token, token)
    admins = configured_admin_emails()
    is_allowed = bool(clean_admin and clean_admin in admins)
    return {
        "authenticated_email": clean_admin,
        "is_admin": is_allowed,
        "admin_email_count": len(admins),
        "admin_emails": admins,
    }


def verify_secret(x_archlens_secret: str) -> None:
    if x_archlens_secret != ARCHLENS_WEBHOOK_SECRET:
        raise HTTPException(status_code=401, detail="Invalid webhook secret")


@app.get("/")
def health_check():
    return {
        "status": "ArchLens API running",
        "storage": "json_file",
        "credits_file": str(CREDITS_FILE),
    }


@app.get("/user/{email}")
def get_user_credits(email: str):
    clean_email = normalise_email(email)
    with STORE_LOCK:
        store = load_store()
        status = store.get("user_meta", {}).get(clean_email, {}).get("status", "active")
    return {
        "email": clean_email,
        "credits": get_balance_from_store(clean_email),
        "status": status,
    }


@app.get("/transactions/{email}")
def get_user_transactions(email: str, x_archlens_secret: str = Header(default="")):
    verify_secret(x_archlens_secret)

    clean_email = normalise_email(email)
    with STORE_LOCK:
        store = load_store()
        txs = [
            tx for tx in store.get("transactions", [])
            if tx.get("email") == clean_email
        ][:50]

    return {
        "email": clean_email,
        "transactions": txs,
    }


@app.post("/internal/user-activity")
def record_user_activity(payload: UserActivityRequest, x_archlens_secret: str = Header(default="")):
    verify_secret(x_archlens_secret)
    clean_email = normalise_email(payload.email)
    if not is_valid_email(clean_email):
        raise HTTPException(status_code=400, detail="Valid email is required")
    with STORE_LOCK:
        store = load_store()
        store.setdefault("users", {}).setdefault(clean_email, 0)
        existing_status = store.get("user_meta", {}).get(clean_email, {}).get("status", "")
        activity_status = payload.status if existing_status not in {"suspended"} else existing_status
        update_user_meta(store, clean_email, payload.plan, activity_status)
        save_store(store)
    return {"success": True, "email": clean_email}


@app.post("/internal/report-generation")
def record_report_generation(payload: ReportRecordRequest, x_archlens_secret: str = Header(default="")):
    verify_secret(x_archlens_secret)
    clean_email = normalise_email(payload.email)
    if not is_valid_email(clean_email):
        raise HTTPException(status_code=400, detail="Valid email is required")
    with STORE_LOCK:
        store = load_store()
        store.setdefault("users", {}).setdefault(clean_email, 0)
        update_user_meta(store, clean_email, payload.plan)
        reports = store.setdefault("reports", [])
        record = {
            "email": clean_email,
            "plan": payload.plan,
            "report_id": str(payload.report_id or ""),
            "project_name": str(payload.project_name or payload.project_address or "Untitled project"),
            "project_address": str(payload.project_address or ""),
            "report_type": str(payload.report_type or "Report"),
            "credits_used": int(payload.credits_used or 0),
            "download_path": str(payload.download_path or ""),
            "status": str(payload.status or "generated"),
            "timestamp": utc_now(),
        }
        reports.insert(0, record)
        store["reports"] = reports[:1000]
        save_store(store)
    return {"success": True, "report": record}


@app.post("/api/wix/add-credits")
def add_credits(payload: CreditRequest, x_archlens_secret: str = Header(default="")):
    verify_secret(x_archlens_secret)

    clean_email = normalise_email(payload.email)
    credits_to_add = int(payload.credits or 0)
    order_id = str(payload.orderId or "").strip()

    if not clean_email:
        raise HTTPException(status_code=400, detail="Email is required")

    if credits_to_add <= 0:
        raise HTTPException(status_code=400, detail="Credits must be greater than zero")

    if not order_id:
        raise HTTPException(status_code=400, detail="Order ID is required")

    with STORE_LOCK:
        store = load_store()
        processed_orders = set(store.get("processed_orders", []))

        # Prevent duplicate credit grants if Wix retries the same paid order event.
        if order_id in processed_orders:
            current_balance = int(store["users"].get(clean_email, 0) or 0)
            return {
                "success": True,
                "duplicate": True,
                "message": "Order already processed. Credits were not added again.",
                "email": clean_email,
                "credits": current_balance,
                "new_balance": current_balance,
                "order_id": order_id,
                "source": payload.source,
            }

        current_balance = int(store["users"].get(clean_email, 0) or 0)
        new_balance = current_balance + credits_to_add

        store["users"][clean_email] = new_balance
        update_user_meta(store, clean_email)
        store.setdefault("processed_orders", []).append(order_id)
        add_transaction(
            store,
            clean_email,
            credits_to_add,
            new_balance,
            reason="credits_added",
            source=payload.source,
            reference=order_id,
        )
        save_store(store)

    print(
        f"ADD CREDITS: {clean_email} +{credits_to_add} credits | "
        f"New balance: {new_balance} | Order: {order_id} | Source: {payload.source}"
    )

    return {
        "success": True,
        "email": clean_email,
        "credits_added": credits_to_add,
        "credits": new_balance,
        "new_balance": new_balance,
        "order_id": order_id,
        "source": payload.source,
    }


@app.post("/deduct-credits")
def deduct_credits(payload: DeductCreditRequest, x_archlens_secret: str = Header(default="")):
    verify_secret(x_archlens_secret)

    clean_email = normalise_email(payload.email)
    credits_to_deduct = int(payload.credits or 0)

    if not clean_email:
        raise HTTPException(status_code=400, detail="Email is required")

    if credits_to_deduct <= 0:
        raise HTTPException(status_code=400, detail="Credits must be greater than zero")

    with STORE_LOCK:
        store = load_store()
        current_balance = int(store["users"].get(clean_email, 0) or 0)
        export_type = str(payload.exportType or "export").strip() or "export"
        report_id = str(payload.reportId or "").strip()
        unlock_key = f"{clean_email}:{report_id}:{export_type}"
        processed_unlocks = set(store.get("processed_unlocks", []))

        if report_id and unlock_key in processed_unlocks:
            return {
                "success": True,
                "duplicate": True,
                "email": clean_email,
                "credits_deducted": 0,
                "credits": current_balance,
                "new_balance": current_balance,
                "report_id": payload.reportId,
                "export_type": payload.exportType,
                "source": payload.source,
            }

        if current_balance < credits_to_deduct:
            raise HTTPException(
                status_code=400,
                detail=f"Not enough credits. Current balance is {current_balance}.",
            )

        new_balance = current_balance - credits_to_deduct
        store["users"][clean_email] = new_balance
        update_user_meta(store, clean_email)
        add_transaction(
            store,
            clean_email,
            -credits_to_deduct,
            new_balance,
            reason=f"unlock_{export_type}",
            source=payload.source,
            reference=payload.reportId,
        )
        if report_id:
            store.setdefault("processed_unlocks", []).append(unlock_key)
            store["processed_unlocks"] = store["processed_unlocks"][-1000:]
        save_store(store)

    print(
        f"DEDUCT CREDITS: {clean_email} -{credits_to_deduct} credits | "
        f"New balance: {new_balance} | Report: {payload.reportId} | "
        f"Export: {payload.exportType} | Source: {payload.source}"
    )

    return {
        "success": True,
        "email": clean_email,
        "credits_deducted": credits_to_deduct,
        "credits": new_balance,
        "new_balance": new_balance,
        "report_id": payload.reportId,
        "export_type": payload.exportType,
        "source": payload.source,
    }


@app.get("/admin/summary")
def admin_summary(
    x_archlens_secret: str = Header(default=""),
    authorization: str = Header(default=""),
    x_archlens_token: str = Header(default=""),
    token: str = "",
):
    verify_admin_access(x_archlens_secret, authorization, x_archlens_token, token)
    with STORE_LOCK:
        store = load_store()
        users = store.get("users", {})
        user_meta = store.get("user_meta", {})
        transactions = store.get("transactions", [])
        reports = store.get("reports", [])
        errors = store.get("errors", [])
        active_plans: Dict[str, int] = {}
        for email in users:
            plan = str(user_meta.get(email, {}).get("plan") or "Unknown")
            active_plans[plan] = active_plans.get(plan, 0) + 1
        credits_used = sum(abs(int(tx.get("amount", 0) or 0)) for tx in transactions if int(tx.get("amount", 0) or 0) < 0)
    return {
        "total_users": len(users),
        "total_reports_generated": len(reports),
        "credits_used": credits_used,
        "active_plans": active_plans,
        "recent_report_generations": reports[:10],
        "recent_credit_changes": transactions[:20],
        "recent_errors": errors[:10],
    }


@app.get("/admin/session")
def admin_session(
    x_archlens_secret: str = Header(default=""),
    authorization: str = Header(default=""),
    x_archlens_token: str = Header(default=""),
    token: str = "",
):
    verify_secret(x_archlens_secret)
    payload = admin_session_payload(authorization, x_archlens_token, token)
    print(
        "ArchLens Admin session:",
        {
            "authenticated_email": payload["authenticated_email"],
            "admin_email_count": payload["admin_email_count"],
            "is_admin": payload["is_admin"],
        },
    )
    if not payload["is_admin"]:
        raise HTTPException(status_code=403, detail="Authenticated user email is not listed in ADMIN_EMAILS")
    return payload


@app.get("/admin/users")
def admin_users(
    x_archlens_secret: str = Header(default=""),
    authorization: str = Header(default=""),
    x_archlens_token: str = Header(default=""),
    token: str = "",
):
    verify_admin_access(x_archlens_secret, authorization, x_archlens_token, token)
    with STORE_LOCK:
        store = load_store()
        users = store.get("users", {})
        user_meta = store.get("user_meta", {})
        reports = store.get("reports", [])
        report_counts: Dict[str, int] = {}
        for report in reports:
            email = normalise_email(report.get("email", ""))
            report_counts[email] = report_counts.get(email, 0) + 1
        rows = []
        for email, credits in sorted(users.items()):
            meta = user_meta.get(email, {})
            rows.append(
                {
                    "email": email,
                    "plan": meta.get("plan", "Unknown"),
                    "credits": int(credits or 0),
                    "last_activity": meta.get("last_activity", ""),
                    "status": meta.get("status", "active"),
                    "reports_generated": report_counts.get(email, 0),
                }
            )
    return {"users": rows}


@app.get("/admin/reports")
def admin_reports(
    x_archlens_secret: str = Header(default=""),
    authorization: str = Header(default=""),
    x_archlens_token: str = Header(default=""),
    token: str = "",
):
    verify_admin_access(x_archlens_secret, authorization, x_archlens_token, token)
    with STORE_LOCK:
        store = load_store()
        reports = store.get("reports", [])[:500]
    return {"reports": reports}


@app.get("/admin/audit-log")
def admin_audit_log(
    x_archlens_secret: str = Header(default=""),
    authorization: str = Header(default=""),
    x_archlens_token: str = Header(default=""),
    token: str = "",
):
    verify_admin_access(x_archlens_secret, authorization, x_archlens_token, token)
    with STORE_LOCK:
        store = load_store()
        audit_log = store.get("audit_log", [])[:500]
    return {"audit_log": audit_log}


@app.post("/admin/credits/adjust")
def admin_adjust_credits(
    payload: AdminAdjustCreditRequest,
    x_archlens_secret: str = Header(default=""),
    authorization: str = Header(default=""),
    x_archlens_token: str = Header(default=""),
    token: str = "",
):
    admin_email = verify_admin_access(x_archlens_secret, authorization, x_archlens_token, token)
    clean_email = normalise_email(payload.email)
    action = str(payload.action or "").strip().lower()
    credits_value = int(payload.credits or 0)
    reason = str(payload.reason or "").strip()

    if not is_valid_email(clean_email):
        raise HTTPException(status_code=400, detail="Valid target email is required")
    if action not in {"add", "remove", "set"}:
        raise HTTPException(status_code=400, detail="Action must be add, remove, or set")
    if credits_value < 0:
        raise HTTPException(status_code=400, detail="Credits must be zero or greater")
    if action in {"add", "remove"} and credits_value == 0:
        raise HTTPException(status_code=400, detail="Credit change cannot be zero")
    if not reason:
        raise HTTPException(status_code=400, detail="Reason is required")

    with STORE_LOCK:
        store = load_store()
        current_balance = int(store.setdefault("users", {}).get(clean_email, 0) or 0)
        if action == "add":
            new_balance = current_balance + credits_value
            change_amount = credits_value
        elif action == "remove":
            new_balance = current_balance - credits_value
            if new_balance < 0:
                raise HTTPException(status_code=400, detail="Credit adjustment would create a negative balance")
            change_amount = -credits_value
        else:
            new_balance = credits_value
            change_amount = new_balance - current_balance

        store["users"][clean_email] = new_balance
        update_user_meta(store, clean_email)
        add_transaction(
            store,
            clean_email,
            change_amount,
            new_balance,
            reason=f"admin_{action}: {reason}",
            source="admin",
            reference=admin_email,
        )
        audit_entry = {
            "admin_email": admin_email,
            "target_user_email": clean_email,
            "previous_credits": current_balance,
            "new_credits": new_balance,
            "change_amount": change_amount,
            "reason": reason,
            "action": action,
            "timestamp": utc_now(),
        }
        audit_log = store.setdefault("audit_log", [])
        audit_log.insert(0, audit_entry)
        store["audit_log"] = audit_log[:1000]
        save_store(store)

    return {
        "success": True,
        "email": clean_email,
        "previous_credits": current_balance,
        "new_credits": new_balance,
        "change_amount": change_amount,
        "reason": reason,
    }


@app.post("/admin/users/status")
def admin_set_user_status(
    payload: AdminUserStatusRequest,
    x_archlens_secret: str = Header(default=""),
    authorization: str = Header(default=""),
    x_archlens_token: str = Header(default=""),
    token: str = "",
):
    admin_email = verify_admin_access(x_archlens_secret, authorization, x_archlens_token, token)
    clean_email = normalise_email(payload.email)
    status = str(payload.status or "").strip().lower()
    reason = str(payload.reason or "").strip()

    if not is_valid_email(clean_email):
        raise HTTPException(status_code=400, detail="Valid target email is required")
    if status not in {"active", "suspended"}:
        raise HTTPException(status_code=400, detail="Status must be active or suspended")
    if not reason:
        raise HTTPException(status_code=400, detail="Reason is required")

    with STORE_LOCK:
        store = load_store()
        store.setdefault("users", {}).setdefault(clean_email, 0)
        meta = store.setdefault("user_meta", {}).get(clean_email, {})
        previous_status = meta.get("status", "active")
        meta["status"] = status
        meta["last_activity"] = utc_now()
        store.setdefault("user_meta", {})[clean_email] = meta
        audit_entry = {
            "admin_email": admin_email,
            "target_user_email": clean_email,
            "previous_status": previous_status,
            "new_status": status,
            "reason": reason,
            "action": "status_update",
            "timestamp": utc_now(),
        }
        audit_log = store.setdefault("audit_log", [])
        audit_log.insert(0, audit_entry)
        store["audit_log"] = audit_log[:1000]
        save_store(store)

    return {
        "success": True,
        "email": clean_email,
        "previous_status": previous_status,
        "status": status,
        "reason": reason,
    }


@app.post("/admin/add-credits")
def admin_add_credits(
    payload: AdminCreditRequest,
    x_archlens_secret: str = Header(default=""),
    authorization: str = Header(default=""),
    x_archlens_token: str = Header(default=""),
    token: str = "",
):
    """
    Protected manual restore/top-up endpoint.
    Use this to restore credits already purchased if a Wix event failed before automation was fixed.
    """
    admin_email = verify_admin_access(x_archlens_secret, authorization, x_archlens_token, token)

    clean_email = normalise_email(payload.email)
    credits_to_add = int(payload.credits or 0)

    if not is_valid_email(clean_email):
        raise HTTPException(status_code=400, detail="Valid email is required")

    if credits_to_add <= 0:
        raise HTTPException(status_code=400, detail="Credits must be greater than zero")

    with STORE_LOCK:
        store = load_store()
        current_balance = int(store["users"].get(clean_email, 0) or 0)
        new_balance = current_balance + credits_to_add

        store["users"][clean_email] = new_balance
        update_user_meta(store, clean_email)
        add_transaction(
            store,
            clean_email,
            credits_to_add,
            new_balance,
            reason=payload.reason,
            source="admin",
            reference=admin_email,
        )
        audit_log = store.setdefault("audit_log", [])
        audit_log.insert(
            0,
            {
                "admin_email": admin_email,
                "target_user_email": clean_email,
                "previous_credits": current_balance,
                "new_credits": new_balance,
                "change_amount": new_balance - current_balance,
                "reason": payload.reason,
                "action": "add",
                "timestamp": utc_now(),
            },
        )
        store["audit_log"] = audit_log[:1000]
        save_store(store)

    return {
        "success": True,
        "email": clean_email,
        "credits_added": credits_to_add,
        "credits": new_balance,
        "new_balance": new_balance,
        "reason": payload.reason,
    }


@app.get("/admin/restore/{email}/{credits}")
def admin_restore_credits(
    email: str,
    credits: int,
    x_archlens_secret: str = Header(default=""),
    authorization: str = Header(default=""),
    x_archlens_token: str = Header(default=""),
    token: str = "",
):
    """
    Protected browser-friendly restore endpoint.
    This requires the x-archlens-secret header, so use Postman/curl for security.
    """
    payload = AdminCreditRequest(email=email, credits=credits, reason="manual_restore")
    return admin_add_credits(
        payload,
        x_archlens_secret=x_archlens_secret,
        authorization=authorization,
        x_archlens_token=x_archlens_token,
        token=token,
    )
