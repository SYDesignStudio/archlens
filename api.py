import os
from fastapi import FastAPI, Header, HTTPException
from pydantic import BaseModel

app = FastAPI()

ARCHLENS_WEBHOOK_SECRET = os.getenv(
    "ARCHLENS_WEBHOOK_SECRET",
    "change-this-secret"
)

# Temporary in-memory credit store for testing only.
# This will reset when the Render service restarts.
# Later, replace this with Supabase/database storage.
USER_CREDITS = {}


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


@app.get("/")
def health_check():
    return {"status": "ArchLens API running"}


@app.get("/user/{email}")
def get_user_credits(email: str):
    clean_email = email.lower().strip()
    return {
        "email": clean_email,
        "credits": USER_CREDITS.get(clean_email, 0)
    }


@app.post("/api/wix/add-credits")
def add_credits(
    payload: CreditRequest,
    x_archlens_secret: str = Header(default="")
):
    if x_archlens_secret != ARCHLENS_WEBHOOK_SECRET:
        raise HTTPException(
            status_code=401,
            detail="Invalid webhook secret"
        )

    if payload.credits <= 0:
        raise HTTPException(
            status_code=400,
            detail="Credits must be greater than zero"
        )

    clean_email = payload.email.lower().strip()

    if not clean_email:
        raise HTTPException(
            status_code=400,
            detail="Email is required"
        )

    USER_CREDITS[clean_email] = USER_CREDITS.get(clean_email, 0) + payload.credits

    print(
        f"ADD CREDITS: {clean_email} "
        f"+{payload.credits} credits | "
        f"New balance: {USER_CREDITS[clean_email]} | "
        f"Order: {payload.orderId} | "
        f"Source: {payload.source}"
    )

    return {
        "success": True,
        "email": clean_email,
        "credits_added": payload.credits,
        "new_balance": USER_CREDITS[clean_email],
        "order_id": payload.orderId,
        "source": payload.source,
    }

@app.post("/deduct-credits")
def deduct_credits(
    payload: DeductCreditRequest,
    x_archlens_secret: str = Header(default="")
):
    if x_archlens_secret != ARCHLENS_WEBHOOK_SECRET:
        raise HTTPException(
            status_code=401,
            detail="Invalid webhook secret"
        )

    clean_email = payload.email.lower().strip()
    credits_to_deduct = int(payload.credits or 0)

    if not clean_email:
        raise HTTPException(
            status_code=400,
            detail="Email is required"
        )

    if credits_to_deduct <= 0:
        raise HTTPException(
            status_code=400,
            detail="Credits must be greater than zero"
        )

    current_balance = USER_CREDITS.get(clean_email, 0)

    if current_balance < credits_to_deduct:
        raise HTTPException(
            status_code=400,
            detail="Not enough credits"
        )

    USER_CREDITS[clean_email] = current_balance - credits_to_deduct

    print(
        f"DEDUCT CREDITS: {clean_email} "
        f"-{credits_to_deduct} credits | "
        f"New balance: {USER_CREDITS[clean_email]} | "
        f"Report: {payload.reportId} | "
        f"Export: {payload.exportType} | "
        f"Source: {payload.source}"
    )

    return {
        "success": True,
        "email": clean_email,
        "credits_deducted": credits_to_deduct,
        "credits": USER_CREDITS[clean_email],
        "new_balance": USER_CREDITS[clean_email],
        "report_id": payload.reportId,
        "export_type": payload.exportType,
        "source": payload.source,
    }

