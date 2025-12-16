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
        data = request.get_json()
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
                        if isinstance(first, dict):
                            return first.get("label") or first.get("value") or ""
                        return first
                    return value
            return ""

        # ---- GLOBAL FILE DETECTOR ----
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
        first_name = (
            get_value("First Name")
            or get_value("First name")
            or get_value("firstname")
        )

        service_type = get_value("What leather service are you interested in?")
        item_type = get_value("What type of leather item?")
        color_selection = get_value("Color Selection")
        customer_email = get_value("Email")

        has_photos = has_any_uploaded_files()

        # ---- CLEAN FIRST NAME ----
        first_name = first_name.strip().title() if isinstance(first_name, str) else ""
        greeting_name = first_name if first_name else "there"

        if not customer_email or not service_type:
            return jsonify({"status": "ignored"}), 200

        token = get_access_token(
            AZURE_TENANT_ID,
            AZURE_CLIENT_ID,
            AZURE_CLIENT_SECRET,
        )

        # ==================================================
        # NO PHOTOS → SHORT EMAIL
        # ==================================================
        if not has_photos:
            email_body = f"""Hi {greeting_name},

Thank you for your interest in ReLeather.

We’d be happy to look into {service_type} for your {item_type}. To provide accurate recommendations and pricing, please send us a few photos, any additional details, and if possible dimensions. We’ll follow up shortly.
"""
            email_body = (
                email_body.replace("\n", "<br/>")
                + "<br/><br/>"
                + OUTLOOK_EMAIL_SIGNATURE
            )

            if token:
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
        # ---- INTRO BLOCK ----
        email_body = f"""Hi {greeting_name},

Thank you for your interest in ReLeather.

Based on the information provided, we recommend our {service_type} for your {item_type}.
"""

        # ---- CONDITIONAL SERVICE BODY ----
        if service_type == "Leather Restoration":
            email_body += """
This service addresses surface wear such as color fading, light scratches, scuffs, stains, and spotting. It also restores the leather’s original uniform color and matte finish. We complete the process with a protective coating to prevent color transfer.

Please note: We cannot repair or restore the grain texture in areas where it has been worn smooth or torn. We can perform minor patching for tears if needed. However, the cosmetic result varies depending on the damage.
"""

        elif service_type == "Leather Cleaning & Conditioning":
            email_body += """
Leather Cleaning removes surface dirt and build up, deep cleans the leather surface. Leather Conditioning moisturizes, softens, strengthens, polishes the leather, and prevents water spotting and cracking. Leather Retouching treats minor scuffs and discoloration, and renews color finish. Leather Protection applies a finish protection.
"""

        elif service_type == "Leather Dyeing (Color Change)":
            email_body += f"""
This service treats the old finish and dyes the leather in your selected color — {color_selection}. It also refreshes the overall finish of the item, enhancing both appearance and longevity. We complete the process with a protective topcoat to prevent color transfer.

Please note: The new surface coating applied during dyeing may reduce the suppleness of the leather. Accent stitching will be dyed to match the new leather color. While we carefully mask fabric strips and linings during restoration, some dye transfer may occur. We take precautions to minimize this.
"""

        elif service_type == "Leather Reupholstery":
            email_body += """
Full Leather Reupholstery replaces all upholstery with new leather of your choice. We offer a wide selection of colors, textures, and finishes. This requires purchasing new leather and disassembly of the upholstery.

Partial Leather Reupholstery service recovers damaged leather for specific cushions. This also requires purchasing new leather and upholstery disassembly.

Please note: We source leather that closely matches your original; however, the worn-in patina of existing leather may not match seamlessly. For accurate measurements and pattern matching, we require the original seat cover for each unique cushion size mailed to us.
"""

        elif service_type == "Foam Replacement & Restuffing":
            email_body += """
This service replaces the seat cushion core with high-resilience (HR) grade foam and adds polyester fiber padding for improved structure and comfort.
"""

        # ---- ENDING BLOCK ----
        email_body += """
Estimated cost: $.

Completion time: 1–2 weeks.

Drop-off: By appointment at our Fullerton, CA shop.

Free Pick Up and Delivery in Orange County.

$200 Pick Up and Delivery in Los Angeles, San Diego, and Riverside County.

Please contact us with any questions or to proceed with your order. Thank you.
"""

        email_body = (
            email_body.replace("\n", "<br/>")
            + "<br/><br/>"
            + OUTLOOK_EMAIL_SIGNATURE
        )

        if token:
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
