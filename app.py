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

        def get_value(name):
            for q in questions:
                if q.get("name") == name:
                    return q.get("value", "")
            return ""

        service_type = get_value("What leather service are you interested in?")
        customer_email = get_value("Email")

        if service_type != "Leather Restoration":
            return jsonify({"status": "ignored"}), 200

        email_subject = "Leather Restoration – ReLeather"

        email_body = """Hi,

Thank you for your interest in ReLeather.

Based on the information provided, we recommend our Leather Restoration for your item.

This service addresses surface wear such as color fading, light scratches, scuffs, stains, and spotting. It also restores the leather’s original uniform color and matte finish. We complete the process with a protective coating to prevent color transfer.

Please note: We cannot repair or restore the grain texture in areas where it has been worn smooth or torn. We can perform minor patching for tears if needed. However, the cosmetic result varies depending on the damage.

Estimated cost: $.

Completion time: 1–2 weeks.

Drop-off: By appointment at our Fullerton, CA shop.

Shipping instructions for non-local customers will be provided after confirming your order.

Please feel free to contact us with any questions or to proceed with your order.
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
                email_subject,
                email_body
            )
        else:
            logging.error("No access token — draft not created")

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
