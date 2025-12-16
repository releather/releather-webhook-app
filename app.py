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
        item_type = get_value("What type of leather item?")
        color_selection = get_value("Color Selection")
        customer_email = get_value("Email")

        if not customer_email or not service_type:
            logging.warning("Missing email or service type")
            return jsonify({"status": "ignored"}), 200

        # ---- ATTACHMENT CHECK ----
        attached_photos = get_files("Attach Photos")
        has_photos = isinstance(attached_photos, list) and len(attached_photos) > 0

        if not has_photos:
            email_body = f"""Hi {customer_first_name},

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

        # ---- INTRO BLOCK ----
        email_body = f"""Hi {customer_first_name},

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

Reference Pricing:
- Standard Sofa Seat Cushion: $1100 each
  Thickness: 4–6"
  Width: 22–26"
  Depth: 20–24"

- Larger Seat Cushions: $1200+ each
  Thickness: 5–8"
  Width: 26–32"
  Depth: 24–34"

Foam core not included. Additional labor cost applies for fixed seating.
"""

        elif service_type == "Foam Replacement & Restuffing":
            email_body += """
This service replaces the seat cushion core with high-resilience (HR) grade foam and adds polyester fiber padding for improved structure and comfort. Foam is available in soft, medium, and firm densities.

Reference Pricing:
- Standard Sofa Seat Cushion: $350–$450 each
- Larger Seat Cushions: $450–$600 each

Please note: We do not use down feather filling. For accurate sizing, we may require the original seat cover mailed to us.
"""

        # ---- ENDING BLOCK ----
        email_body += """
Estimated cost: $.

Completion time: 1–2 weeks.

Drop-off: By appointment at our Fullerton, CA shop.

Free Pick Up and Delivery in Orange County.

$200 Pick Up and Delivery in Los Angeles, San Diego, and Riverside County.

Non-local customers: Shipping instructions for mailed-in orders will be provided after confirming your order. Return shipping is quoted separately.

Please contact us with any questions or to proceed with your order. Thank you.
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
