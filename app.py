import os
import logging
import requests
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

    logging.info(f"Token status: {response.status_code}")

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
        "body": {
            "contentType": "HTML",
            "content": body,
        },
        "toRecipients": [
            {"emailAddress": {"address": recipient_email}}
        ],
        "isDraft": True,
    }

    response = requests.post(url, headers=headers, json=payload, timeout=10)

    if response.status_code not in (200, 201):
        logging.error(f"Graph error: {response.status_code} {response.text}")
        return False

    logging.info("Outlook draft created successfully")
    return True


# ---- WEBHOOK ----
@app.route("/webhook", methods=["POST"])
def webhook():
    try:
        data = request.get_json()
        logging.info("Webhook received")

        questions = data.get("submission", {}).get("questions", [])

        # ---- VALUE EXTRACTOR ----
        def get_value(name):
            for q in questions:
                if q.get("name") == name:
                    value = q.get("value", "")
                    if isinstance(value, list):
                        if not value:
                            return ""
                        first = value[0]
                        if isinstance(first, str):
                            return first
                        if isinstance(first, dict):
                            return first.get("label") or first.get("value") or ""
                    return value
            return ""

        # ---- ✅ GLOBAL FILE DETECTOR (FIX) ----
        def has_any_uploaded_files(questions):
            for q in questions:
                value = q.get("value")
                if isinstance(value, list):
                    for f in value:
                        if isinstance(f, dict):
                            if f.get("url") or f.get("filename") or f.get("name"):
                                return True
            return False

        has_photos = has_any_uploaded_files(questions)

        # ---- FIRST NAME (ROBUST) ----
        first_name = (
            get_value("First Name")
            or get_value("First name")
            or get_value("first_name")
            or get_value("firstname")
        )

        # ---- FORM VALUES ----
        service_type = get_value("What leather service are you interested in?")
        item_type = get_value("What type of leather item?")
        color_selection = get_value("Color Selection")
        customer_email = get_value("Email")

        # ---- CLEAN FIRST NAME ----
        if isinstance(first_name, str):
            first_name = first_name.strip().title()
        else:
            first_name = ""

        greeting_name = first_name if first_name else "there"

        if not customer_email or not service_type:
            logging.warning("Missing email or service type")
            return jsonify({"status": "ignored"}), 200

        # ==================================================
        # ✅ NO PHOTOS → SHORT EMAIL
        # ==================================================
        if not has_photos:
            email_body = f"""Hi {greeting_name},

Thank you for your interest in ReLeather.

We’d be happy to look into {service_type} for your {item_type}. To provide accurate recommendations and pricing, please send us a few photos, any additional details, and if possible dimensions. We’ll follow up shortly.
"""

            email_body = email_body.replace("\n", "<br/>") + "<br/><br/>" + OUTLOOK_EMAIL_SIGNATURE

            access_token = get_access_token(
                AZURE_TENANT_ID,
                AZURE_CLIENT_ID,
                AZURE_CLIENT_SECRET
            )

            if access_token:
                create_outlook_draft(
                    access_token,
                    OUTLOOK_SENDER_EMAIL,
                    customer_email,
                    f"{service_type} – ReLeather",
                    email_body
                )

            return jsonify({"status": "awaiting_photos"}), 200

        # ==================================================
        # ✅ PHOTOS PRESENT → FULL EMAIL
        # ==================================================
        email_body = f"""Hi {greeting_name},

Thank you for your interest in ReLeather.

Based on the information provided, we recommend our {service_type} for your {item_type}.
"""

        if service_type == "Leather Restoration":
            email_body += """
This service addresses surface wear such as color fading, light scratches, scuffs, stains, and spotting. It also restores the leather’s original uniform color and matte finish.
"""

        elif service_type == "Leather Cleaning & Conditioning":
            email_body += """
Leather Cleaning removes surface dirt and build up, deep cleans the leather surface. Leather Conditioning moisturizes and protects the leather.
"""

        elif service_type == "Leather Dyeing (Color Change)":
            email_body += f"""
This service treats the old finish and dyes the leather in your selected color — {color_selection}. We complete the process with a protective topcoat.
"""

        elif service_type == "Leather Reupholstery":
            email_body += """
Full Leather Reupholstery replaces all upholstery with new leather of your choice.
"""

        elif service_type == "Foam Replacement & Restuffing":
            email_body += """
This service replaces the seat cushion core with high-resilience foam.
"""

        email_body += """
Estimated cost: $.

Completion time: 1–2 weeks.

Drop-off: By appointment at our Fullerton, CA shop.

Please contact us with any questions or to proceed.
"""

        email_body = email_body.replace("\n", "<br/>") + "<br/><br/>" + OUTLOOK_EMAIL_SIGNATURE

        access_token = get_access_token(
            AZURE_TENANT_ID,
            AZURE_CLIENT_ID,
            AZURE_CLIENT_SECRET
        )

        if access_token:
            create_outlook_draft(
                access_token,
                OUTLOOK_SENDER_EMAIL,
                customer_email,
                f"{service_type} – ReLeather",
                email_body
            )

        return jsonify({"status": "processed"}), 200

    except Exception as e:
        logging.error(str(e))
        return jsonify({"error": "internal error"}), 500


# ---- HEALTH CHECK ----
@app.route("/", methods=["GET"])
def index():
    return "Webhook server is running."


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
