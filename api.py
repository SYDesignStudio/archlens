import json
import os
import threading
from pathlib import Path
from typing import Dict, List

from fastapi import FastAPI, Header, HTTPException
from pydantic import BaseModel

app = FastAPI(title="ArchLens Credits API")

ARCHLENS_WEBHOOK_SECRET = os.getenv(
    "ARCHLENS_WEBHOOK_SECRET",
    "archlens_secure_2026_SYDS_92838"
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


def normalise_email(email: str) -> str:
    return str(email or "").strip().lower()


def default_store() -> Dict:
    return {
        "users": {},
        "processed_orders": [],
        "transactions": [],
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
    data.setdefault("transactions", [])
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
        },
    )
    store["transactions"] = transactions[:500]


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
    return {
        "email": clean_email,
        "credits": get_balance_from_store(clean_email),
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

        if current_balance < credits_to_deduct:
            raise HTTPException(
                status_code=400,
                detail=f"Not enough credits. Current balance is {current_balance}.",
            )

        new_balance = current_balance - credits_to_deduct
        store["users"][clean_email] = new_balance
        add_transaction(
            store,
            clean_email,
            -credits_to_deduct,
            new_balance,
            reason=f"unlock_{payload.exportType or 'export'}",
            source=payload.source,
            reference=payload.reportId,
        )
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


@app.post("/admin/add-credits")
def admin_add_credits(payload: AdminCreditRequest, x_archlens_secret: str = Header(default="")):
    """
    Protected manual restore/top-up endpoint.
    Use this to restore credits already purchased if a Wix event failed before automation was fixed.
    """
    verify_secret(x_archlens_secret)

    clean_email = normalise_email(payload.email)
    credits_to_add = int(payload.credits or 0)

    if not clean_email:
        raise HTTPException(status_code=400, detail="Email is required")

    if credits_to_add == 0:
        raise HTTPException(status_code=400, detail="Credits cannot be zero")

    with STORE_LOCK:
        store = load_store()
        current_balance = int(store["users"].get(clean_email, 0) or 0)
        new_balance = max(0, current_balance + credits_to_add)

        store["users"][clean_email] = new_balance
        add_transaction(
            store,
            clean_email,
            credits_to_add,
            new_balance,
            reason=payload.reason,
            source="admin",
            reference="manual",
        )
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
def admin_restore_credits(email: str, credits: int, x_archlens_secret: str = Header(default="")):
    """
    Protected browser-friendly restore endpoint.
    This requires the x-archlens-secret header, so use Postman/curl for security.
    """
    payload = AdminCreditRequest(email=email, credits=credits, reason="manual_restore")
    return admin_add_credits(payload, x_archlens_secret=x_archlens_secret)
