import os
import logging
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

# ---- REQUIRED IMPORTS (must already exist in your project) ----
from outlook import get_access_token, create_outlook_draft


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

        return jsonify({"status": "draft_created"}), 200

    except Exception as e:
        logging.error(str(e))
        return jsonify({"error": "internal error"}), 500


@app.route("/", methods=["GET"])
def index():
    return "Webhook server is running."


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
