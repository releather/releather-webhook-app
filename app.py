import os
import logging
import requests
import json
from flask import Flask, request, jsonify
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

# ---- ENV VARS ----
AZURE_TENANT_ID = os.environ.get("AZURE_TENANT_ID")
AZURE_CLIENT_ID = os.environ.get("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")
OUTLOOK_SENDER_EMAIL = os.environ.get("OUTLOOK_SENDER_EMAIL")
OUTLOOK_EMAIL_SIGNATURE = os.environ.get("OUTLOOK_EMAIL_SIGNATURE", "")

# ---- MICROSOFT GRAPH HELPERS ----
def get_access_token(tenant_id, client_id, client_secret):
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    response = requests.post(
        token_url,
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=10,
    )

    if response.status_code != 200:
        logging.error(f"Token error: {response.text}")
        return None

    return response.json().get("access_token")


def create_outlook_draft(access_token, sender_email, recipient_email, subject, body):
    url = f"https://graph.microsoft.com/v1.0/users/{sender_email}/messages"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    payload = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": body},
        "toRecipients": [{"emailAddress": {"address": recipient_email}}],
        "isDraft": True,
    }

    response = requests.post(url, headers=headers, json=payload, timeout=10)

    if response.status_code not in (200, 201):
        logging.error(f"Graph error: {response.status_code} {response.text}")
        return False

    return True


# ---- WEBHOOK ----
@app.route("/webhook", methods=["POST"])
def webhook():
    try:
        data = request.get_json() or {}

        logging.info(json.dumps(data, indent=2))

        questions = data.get("submission", {}).get("questions", [])

        # ---- SAFE VALUE EXTRACTOR ----
        def get_value(name):
            for q in questions:
                if q.get("name") == name:
                    value = q.get("value")

                    if value is None:
                        return ""

                    if isinstance(value, list):
                        if not value:
                            return ""
                        first = value[0]
                        if isinstance(first, dict):
                            return first.get("label") or first.get("value") or ""
                        return first

                    return value
            return ""

        # ---- SAFE STR ----
        def safe_text(value):
            if value is None:
                return ""
            return str(value).strip()

        # ---- FILE DETECTOR ----
        def has_any_uploaded_files():
            for q in questions:
                value = q.get("value")
                if isinstance(value, list):
                    for f in value:
                        if isinstance(f, dict) and (
                            f.get("url") or f.get("filename") or f.get("name")
                        ):
                            return True
            return False

        # ---- FORM VALUES ----
        first_name = safe_text(
            get_value("First Name")
            or get_value("First name")
            or get_value("firstname")
        )

        service_type = safe_text(get_value("What leather service are you interested in?"))
        item_type = safe_text(get_value("What type of leather item?"))
        color_selection = get_value("Color Selection")
        customer_email = safe_text(get_value("Email"))

        has_photos = has_any_uploaded_files()

        greeting_name = first_name.title() if first_name else "there"

        # ---- DEBUG LOGGING (IMPORTANT) ----
        logging.info(f"[WEBHOOK CHECK] email={customer_email} service={service_type}")

        # ---- HARD DEBUG SAFETY ----
        if not customer_email:
            logging.error("Missing email field - cannot send email draft")

        if not service_type:
            logging.error("Missing service_type - continuing with fallback value")
            service_type = "Unknown Service"

        token = get_access_token(
            AZURE_TENANT_ID,
            AZURE_CLIENT_ID,
            AZURE_CLIENT_SECRET,
        )

        if not token:
            logging.error("Failed to get Microsoft Graph token")

        # ==================================================
        # NO PHOTOS → SHORT EMAIL
        # ==================================================
        if not has_photos:
            email_body = f"""Hi {greeting_name},

Thank you for your interest in ReLeather.

We’d be happy to look into {service_type} for your {item_type}. Please send photos for accurate pricing.
"""

            email_body = (
                email_body.replace("\n", "<br/>")
                + "<br/><br/>"
                + OUTLOOK_EMAIL_SIGNATURE
            )

            if token and customer_email:
                create_outlook_draft(
                    token,
                    OUTLOOK_SENDER_EMAIL,
                    customer_email,
                    f"{service_type} – ReLeather",
                    email_body,
                )

            return jsonify({"status": "awaiting_photos"}), 200

        # ==================================================
        # PHOTOS PRESENT → FULL EMAIL
        # ==================================================
        email_body = f"""Hi {greeting_name},

Thank you for your interest in ReLeather.

Based on your submission, we recommend our {service_type} for your {item_type}.
"""

        if service_type == "Leather Restoration":
            email_body += """
This service restores color, removes wear, and applies protective coating.
https://www.releather.com/services#leather-restoration
"""

        elif service_type == "Leather Cleaning & Conditioning":
            email_body += """
This service deep cleans, conditions, and restores leather softness.
https://www.releather.com/services#leather-cleaning
"""

        elif service_type == "Leather Dyeing (Color Change)":
            email_body += f"""
We dye leather into your selected color: {color_selection}
https://www.releather.com/services#leather-dyeing
"""

        elif service_type == "Leather Reupholstery":
            email_body += """
Full replacement of leather upholstery with new materials.
https://www.releather.com/services#leather-upholstery
"""

        elif service_type == "Foam Replacement & Restuffing":
            email_body += """
Refilling cushions with high-density foam and fiber support.
https://www.releather.com/services#foamrestuff
"""

        email_body += """
Estimated cost: $.
Completion time: 2–4 weeks.

Drop-off: 751 S State College Unit 38, Fullerton, CA 92831.

Thank you.
"""

        email_body = (
            email_body.replace("\n", "<br/>")
            + "<br/><br/>"
            + OUTLOOK_EMAIL_SIGNATURE
        )

        if token and customer_email:
            create_outlook_draft(
                token,
                OUTLOOK_SENDER_EMAIL,
                customer_email,
                f"{service_type} – ReLeather",
                email_body,
            )

        return jsonify({"status": "draft_created"}), 200

    except Exception as e:
        logging.exception("Webhook error")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/", methods=["GET"])
def index():
    return "Webhook server is running."
