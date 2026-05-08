import os
from fastapi import FastAPI, Header, HTTPException
from pydantic import BaseModel

app = FastAPI()

ARCHLENS_WEBHOOK_SECRET = os.getenv(
    "ARCHLENS_WEBHOOK_SECRET",
    "change-this-secret"
)

class CreditRequest(BaseModel):
    email: str
    credits: int
    orderId: str
    source: str = "wix_store"


@app.get("/")
def health_check():
    return {"status": "ArchLens API running"}


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

    # TEMP TEST VERSION
    # Later this will update Supabase/database

    print(
        f"ADD CREDITS: {payload.email} "
        f"+{payload.credits} credits | "
        f"Order: {payload.orderId}"
    )

    return {
        "success": True,
        "email": payload.email,
        "credits_added": payload.credits,
        "order_id": payload.orderId,
    }
