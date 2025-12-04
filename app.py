import os
import logging
import re
import requests
from flask import Flask, request, jsonify
from dotenv import load_dotenv

# Google Gemini SDK
import google.generativeai as genai
from google.generativeai import GenerationConfig

# Load environment variables
load_dotenv()

# Create Flask app
app = Flask(__name__)

# Logging setup
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s in %(module)s: %(message)s'
)

# -----------------------------------------------------
#  GOOGLE GEMINI CONFIGURATION
# -----------------------------------------------------

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

if not GEMINI_API_KEY:
    logging.error("GEMINI_API_KEY not found in environment variables")

# Configure API key
genai.configure(api_key=GEMINI_API_KEY)

# Generation config
generation_config = GenerationConfig(
    temperature=0.7,
    max_output_tokens=512
)

# IMPORTANT:
# Render is installing an older google-generativeai version.
# That version does NOT support "-002". It ONLY supports "gemini-1.5-flash".
model = genai.GenerativeModel(
    model_name="gemini-1.5-flash",
    generation_config=generation_config
)

# -----------------------------------------------------
#  HELPER FUNCTIONS
# -----------------------------------------------------

def sanitize_text(text):
    """Remove escape characters & clean up Outlook formatting."""
    if not text:
        return ""
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def ask_gemini(prompt):
    """Send a clean prompt to Gemini and return the clean response."""
    try:
        clean_prompt = sanitize_text(prompt)
        response = model.generate_content(clean_prompt)
        return sanitize_text(response.text)
    except Exception as e:
        logging.error(f"Gemini error: {e}")
        return "There was an error processing your request."

def send_email_reply(access_token, thread_id, reply_text):
    """Send an email reply through Microsoft Graph."""
    try:
        url = f"https://graph.microsoft.com/v1.0/me/messages/{thread_id}/reply"

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        body = {
            "comment": reply_text
        }

        r = requests.post(url, json=body, headers=headers)

        if r.status_code >= 200 and r.status_code < 300:
            logging.info("Reply sent successfully.")
            return True
        else:
            logging.error(f"Failed to send reply: {r.text}")
            return False

    except Exception as e:
        logging.error(f"Error sending email reply: {e}")
        return False

# --- Flask Webhook Route ---
@app.route('/webhook', methods=['POST'])
def webhook():
    try:
        data = request.get_json()
        logging.info(f"Received webhook data: {data}")

        if not data or 'submission' not in data or 'questions' not in data.get('submission', {}):
            logging.warning("Invalid webhook data received: Missing 'submission' or 'questions' key.")
            return jsonify({"error": "Invalid webhook data"}), 400

        submission = data.get('submission', {})
        questions = submission.get('questions', [])

        customer_first_name = get_question_value(questions, 'First Name')
        customer_email = get_question_value(questions, 'Email')
        service_type = get_question_value(questions, 'What leather service are you interested in?')
        item_type = get_question_value(questions, 'What type of leather item?')
        color_selection = get_question_value(questions, 'Color Selection')
        cushions_detachable_fixed = get_question_value(questions, 'Are the seat cushions detachable or fixed to the furniture?')
        how_many_sofas = get_question_value(questions, 'How many sofas?', '0')
        how_many_chairs = get_question_value(questions, 'How many chairs?', '0')
        how_many_cushions = get_question_value(questions, 'How many cushions?', '0')
        share_further_details = get_question_value(questions, 'Share further details:')

        address_field_value = get_question_value(questions, 'Untitled Address field', {})
        zip_code_str = address_field_value.get('zipCode', '') if isinstance(address_field_value, dict) else ""
        state = address_field_value.get('state', '') if isinstance(address_field_value, dict) else ""

        num_sofas = int(how_many_sofas) if how_many_sofas and how_many_sofas.isdigit() else 0
        num_chairs = int(how_many_chairs) if how_many_chairs and how_many_chairs.isdigit() else 0
        num_cushions = int(how_many_cushions) if how_many_cushions and how_many_cushions.isdigit() else 0

        file_upload_questions_values = [
            get_question_value(questions, 'Upload a file (1)'), get_question_value(questions, 'Upload a file (2)'),
            get_question_value(questions, 'Upload a file (3)'), get_question_value(questions, 'Upload a file (5)'),
            get_question_value(questions, 'Upload a file'), get_question_value(questions, 'Upload a file (4)')
        ]
        any_files_uploaded = any(value.strip() != "" for value in file_upload_questions_values)

        plural_item_type = item_type
        if item_type:
            total_explicit_items = num_sofas + num_chairs + num_cushions
            if total_explicit_items > 1:
                if item_type.lower() == 'sofa': plural_item_type = 'sofas'
                elif item_type.lower() == 'chair': plural_item_type = 'chairs'
                elif item_type.lower() == 'cushion': plural_item_type = 'cushions'
                elif item_type.lower() == 'bag': plural_item_type = 'bags'
                elif item_type.lower() == 'coat': plural_item_type = 'coats'
                elif item_type.lower() == 'car': plural_item_type = 'cars'
                else: plural_item_type = item_type + 's'
            elif total_explicit_items == 1:
                plural_item_type = item_type
            else:
                if item_type.lower() in ["bag", "coat", "car"]:
                    plural_item_type = item_type + 's'
                else:
                    plural_item_type = item_type

        zip_code = None
        try:
            if zip_code_str:
                zip_code = int(zip_code_str)
        except (ValueError, TypeError):
            pass

        is_out_of_state = False
        if zip_code is not None:
            if not (90000 <= zip_code <= 92899):
                is_out_of_state = True

        county_name = ""
        if zip_code is not None:
            if (90000 <= zip_code <= 90299) or (90700 <= zip_code <= 90899) or (91000 <= zip_code <= 91899) or (93500 <= zip_code <= 93599):
                county_name = "Los Angeles County"
            elif 92000 <= zip_code <= 92199:
                county_name = "San Diego County"
            elif 92500 <= zip_code <= 92599:
                county_name = "Riverside County"
            elif (92300 <= zip_code <= 92499) or (92857 <= zip_code <= 92859):
                county_name = "San Bernardino County"
            elif (93000 <= zip_code <= 93099):
                county_name = "Ventura County"
            elif (92600 <= zip_code <= 92899):
                county_name = "Orange County"
            elif (92200 <= zip_code <= 92299):
                 county_name = "Imperial County"
            elif (93100 <= zip_code <= 93199):
                county_name = "Santa Barbara County"

        full_text_data = f"{service_type or ''} {item_type or ''} {share_further_details or ''}".lower()

        # --- CRITICAL RULES (Python-side processing - "Output Only" conditions) ---
        # These rules will cause the function to return immediately if matched.
        # This prevents the main Gemini prompt from being sent.

        # 1. "bonded" keyword condition
        if "bonded" in full_text_data:
            logging.info("Triggered 'bonded' keyword critical rule.")
            response_output = f"""Hi {customer_first_name},

Thank you for your interest in ReLeather.

Based on the pictures, your furniture appears to be upholstered in bonded, faux, or split leather. Unfortunately, we do not treat, repair, or reupholster this type of material. If you’re looking for more information about bonded leather, check out our informational blog post: https://www.releather.com/what-is-bonded-leather

However, we’re happy to offer the following options that may support your project:

• If you are sourcing leather for your own furniture project, we recommend visiting [https://www.releather.com/leather-for-upholstery](https://www.releather.com/leather-for-upholstery) for a wide selection of premium hides.
• If you're considering replacing your piece, you can explore our collection of American-made quality leather furniture at: [leathera.com/furniture](https://leathera.com/furniture)
• For ongoing care of your leather, we offer a gentle, high-quality leather conditioner suitable for most leather types. You can find it here: [releather.com/leather-conditioner](https://releather.com/leather-conditioner)"""
            email_subject = "ReLeather Inquiry - Bonded/Faux Leather"
            # Return immediately for critical rules
            access_token = get_access_token(AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET)
            if access_token:
                draft_created = create_outlook_draft(access_token, OUTLOOK_SENDER_EMAIL, customer_email, email_subject, response_output.replace('\n', '<br/>') + OUTLOOK_EMAIL_SIGNATURE)
                if draft_created:
                    logging.info(f"Outlook draft successfully created for {customer_email}.")
                else:
                    logging.error(f"Failed to create Outlook draft for {customer_email}. Check Graph API configuration and permissions.")
            else:
                logging.error("Could not obtain access token for Microsoft Graph API. Outlook draft not created.")
            return jsonify({"response": response_output, "subject": email_subject, "customer_email": customer_email})

        # 2. "low-grade" keyword condition
        if "low-grade" in full_text_data:
            logging.info("Triggered 'low-grade' keyword critical rule.")
            response_output = f"""Hi {customer_first_name},

Thank you for your interest in ReLeather.

Based on your photos and the information provided, this type of leather has a surface-level color coating rather than being dyed through. Our color restoration process is limited to reapplying color as it was originally manufactured. While the treatment is surface-level, we apply a thorough foundation layer to improve durability. The leather will still wear according to the natural characteristics of this material. As a result, we cannot offer our service guarantee for this type of leather.

As an alternative, we can offer Leather Reupholstery. Our Full Leather Upholstery service replacing all upholstery with new leather of your choice. We offer a wide selection of colors, textures, and finishes. If only certain cushions or sections are being replaced, our Partial Leather Upholstery service recovers damaged leather with closely matching leather. This requires purchasing new leather and disassembly of the upholstery.

Please contact if you would like to move forward with an appointment. Thank you."""
            email_subject = "ReLeather Inquiry - Low-Grade Leather"
            access_token = get_access_token(AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET)
            if access_token:
                draft_created = create_outlook_draft(access_token, OUTLOOK_SENDER_EMAIL, customer_email, email_subject, response_output.replace('\n', '<br/>') + OUTLOOK_EMAIL_SIGNATURE)
                if draft_created:
                    logging.info(f"Outlook draft successfully created for {customer_email}.")
                else:
                    logging.error(f"Failed to create Outlook draft for {customer_email}. Check Graph API configuration and permissions.")
            else:
                logging.error("Could not obtain access token for Microsoft Graph API. Outlook draft not created.")
            return jsonify({"response": response_output, "subject": email_subject, "customer_email": customer_email})

        # 3. No files uploaded condition (critical)
        if not any_files_uploaded:
            logging.info("Triggered 'no files uploaded' critical rule.")
            response_output = f"""Hi {customer_first_name},

Thank you for your interest in ReLeather.

We’d be happy to look into {service_type} for your {item_type}. To provide accurate recommendations and pricing, please send us a few photos and any additional details you can share. If available, you may also share the dimensions. We’ll follow up as soon as we receive them.

Thank you!"""
            email_subject = f"ReLeather Inquiry for {item_type}"
            access_token = get_access_token(AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET)
            if access_token:
                draft_created = create_outlook_draft(access_token, OUTLOOK_SENDER_EMAIL, customer_email, email_subject, response_output.replace('\n', '<br/>') + OUTLOOK_EMAIL_SIGNATURE)
                if draft_created:
                    logging.info(f"Outlook draft successfully created for {customer_email}.")
                else:
                    logging.error(f"Failed to create Outlook draft for {customer_email}. Check Graph API configuration and permissions.")
            else:
                logging.error("Could not obtain access token for Microsoft Graph API. Outlook draft not created.")
            return jsonify({"response": response_output, "subject": email_subject, "customer_email": customer_email})

        # 4. Out-of-state fixed furniture condition (critical)
        if service_type in ["Leather Restoration", "Leather Cleaning & Conditioning", "Leather Dyeing (Color Change)", "Foam Replacement"] and \
           item_type in ["Sofa", "Chair", "Cushion"] and \
           cushions_detachable_fixed == "Fixed" and is_out_of_state:
            logging.info("Triggered 'out-of-state fixed furniture' critical rule.")
            response_output = f"""Hi {customer_first_name},

Thank you for your interest in ReLeather.

Please note that we are located in Southern California, and unfortunately, large and fixed seating upholstery services are limited to our local service area. Due to the size and logistics involved, we are unable to accommodate these projects.

However, we’re happy to offer the following options that may support your project:

• If you are sourcing leather for your own furniture project, we recommend visiting [https://www.releather.com/leather-for-upholstery](https://www.releather.com/leather-for-upholstery) for a wide selection of premium hides.
• If you're considering replacing your piece, you can explore our collection of American-made quality leather furniture at: [leathera.com/furniture](https://leathera.com/furniture)
• For ongoing care of your leather, we offer a gentle, high-quality leather conditioner suitable for most leather types. You can find it here: [releather.com/leather-conditioner](https://releather.com/leather-conditioner)"""
            email_subject = "ReLeather Inquiry - Out-of-State Service"
            access_token = get_access_token(AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET)
            if access_token:
                draft_created = create_outlook_draft(access_token, OUTLOOK_SENDER_EMAIL, customer_email, email_subject, response_output.replace('\n', '<br/>') + OUTLOOK_EMAIL_SIGNATURE)
                if draft_created:
                    logging.info(f"Outlook draft successfully created for {customer_email}.")
                else:
                    logging.error(f"Failed to create Outlook draft for {customer_email}. Check Graph API configuration and permissions.")
            else:
                logging.error("Could not obtain access token for Microsoft Graph API. Outlook draft not created.")
            return jsonify({"response": response_output, "subject": email_subject, "customer_email": customer_email})

        # 5. Out-of-state detachable furniture condition (critical)
        if service_type in ["Leather Reupholstery", "Foam Replacement & Restuffing"] and \
           item_type in ["Sofa", "Chair"] and \
           cushions_detachable_fixed == "Detachable" and is_out_of_state:
            logging.info("Triggered 'out-of-state detachable furniture' critical rule.")
            response_output = f"""Hi {customer_first_name},

Thank you for your interest in ReLeather.

Please note that we are located in Southern California, and unfortunately, large and fixed seating upholstery services are limited to our local service area. Due to the size and logistics involved, we are unable to accommodate these projects.

However, we’re happy to offer the following options that may support your project:

• If you are sourcing leather for your own furniture project, we recommend visiting [https://www.releather.com/leather-for-upholstery](https://www.releather.com/leather-for-upholstery) for a wide selection of premium hides.
• If you're considering replacing your piece, you can explore our collection of American-made quality leather furniture at: [leathera.com/furniture](https://leathera.com/furniture)
• For ongoing care of your leather, we offer a gentle, high-quality leather conditioner suitable for most leather types. You can find it here: [releather.com/leather-conditioner](https://releather.com/leather-conditioner)"""
            email_subject = "ReLeather Inquiry - Out-of-State Service"
            access_token = get_access_token(AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET)
            if access_token:
                draft_created = create_outlook_draft(access_token, OUTLOOK_SENDER_EMAIL, customer_email, email_subject, response_output.replace('\n', '<br/>') + OUTLOOK_EMAIL_SIGNATURE)
                if draft_created:
                    logging.info(f"Outlook draft successfully created for {customer_email}.")
                else:
                    logging.error(f"Failed to create Outlook draft for {customer_email}. Check Graph API configuration and permissions.")
            else:
                logging.error("Could not obtain access token for Microsoft Graph API. Outlook draft not created.")
            return jsonify({"response": response_output, "subject": email_subject, "customer_email": customer_email})

        # 6. General out-of-state condition for specific items (critical)
        if service_type in ["Leather Restoration", "Leather Cleaning & Conditioning", "Leather Dyeing (Color Change)", "Leather Reupholstery"] and \
           item_type in ["Car", "Sofa", "Chair"] and is_out_of_state:
            logging.info("Triggered 'general out-of-state' critical rule.")
            response_output = f"""Hi {customer_first_name},

Thank you for your interest in ReLeather.

Please note that we are located in Southern California, and unfortunately, we are unable to accommodate projects out of our service area.

However, we’re happy to offer the following options that may support your project:

• If you are sourcing leather for your own project, we recommend visiting [https://www.releather.com/leather-for-upholstery](https://www.releather.com/leather-for-upholstery) for a wide selection of premium hides.
• For ongoing care of your leather, we offer a gentle, high-quality leather conditioner suitable for most leather types. You can find it here: [releather.com/leather-conditioner](https://releather.com/leather-conditioner)"""
            email_subject = "ReLeather Inquiry - Out-of-State Service"
            access_token = get_access_token(AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET)
            if access_token:
                draft_created = create_outlook_draft(access_token, OUTLOOK_SENDER_EMAIL, customer_email, email_subject, response_output.replace('\n', '<br/>') + OUTLOOK_EMAIL_SIGNATURE)
                if draft_created:
                    logging.info(f"Outlook draft successfully created for {customer_email}.")
                else:
                    logging.error(f"Failed to create Outlook draft for {customer_email}. Check Graph API configuration and permissions.")
            else:
                logging.error("Could not obtain access token for Microsoft Graph API. Outlook draft not created.")
            return jsonify({"response": response_output, "subject": email_subject, "customer_email": customer_email})

        # 7. Car dyeing/reupholstery condition (critical)
        if service_type in ["Leather Dyeing (Color Change)", "Leather Reupholstery"] and item_type == "Car":
            logging.info("Triggered 'car dyeing/reupholstery' critical rule.")
            response_output = f"""Hi {customer_first_name},

Thank you for your interest in ReLeather.

However, we do not offer {service_type} services for {item_type}s. We restore and re-dye car seats to their original color. For more information, please visit: https://www.releather.com/auto-leather-dyeing"""
            email_subject = f"ReLeather Inquiry - {service_type} for {item_type}"
            access_token = get_access_token(AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET)
            if access_token:
                draft_created = create_outlook_draft(access_token, OUTLOOK_SENDER_EMAIL, customer_email, email_subject, response_output.replace('\n', '<br/>') + OUTLOOK_EMAIL_SIGNATURE)
                if draft_created:
                    logging.info(f"Outlook draft successfully created for {customer_email}.")
                else:
                    logging.error(f"Failed to create Outlook draft for {customer_email}. Check Graph API configuration and permissions.")
            else:
                logging.error("Could not obtain access token for Microsoft Graph API. Outlook draft not created.")
            return jsonify({"response": response_output, "subject": email_subject, "customer_email": customer_email})

        # 8. Bag/jacket reupholstery condition (critical)
        if service_type == "Leather Reupholstery" and item_type in ["Bag", "Coat"]:
            logging.info("Triggered 'bag/jacket reupholstery' critical rule.")
            response_output = f"""Hi {customer_first_name},

Thank thank you for your interest in ReLeather.

However, we do not offer {service_type} services for {item_type}s. We restore and re-dye {item_type}s to their original color. For more information, please visit: https://www.releather.com/gallery/leather-redyeing-handbag and https://www.releather.com/leather-restoration-jackets-coats"""
            email_subject = f"ReLeather Inquiry - {service_type} for {item_type}"
            access_token = get_access_token(AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET)
            if access_token:
                draft_created = create_outlook_draft(access_token, OUTLOOK_SENDER_EMAIL, customer_email, email_subject, response_output.replace('\n', '<br/>') + OUTLOOK_EMAIL_SIGNATURE)
                if draft_created:
                    logging.info(f"Outlook draft successfully created for {customer_email}.")
                else:
                    logging.error(f"Failed to create Outlook draft for {customer_email}. Check Graph API configuration and permissions.")
            else:
                logging.error("Could not obtain access token for Microsoft Graph API. Outlook draft not created.")
            return jsonify({"response": response_output, "subject": email_subject, "customer_email": customer_email})

        # --- Initialize email block content based on form data (Python-side logic) ---
        # These will be populated by Python and then inserted into the final Gemini prompt.
        explanation_block = ""
        disclaimer_block = ""
        estimate_block = ""
        completion_block = ""
        delivery_block = ""

        # --- Populate Explanation Block ---
        if service_type == "Leather Restoration":
            explanation_block = "This service addresses surface wear such as color fading, light scratches, scuffs, stains, and spotting. It also restores the leather’s original uniform color and matte finish. We complete the process with a protective coating to prevent color transfer."
        elif service_type == "Leather Cleaning & Conditioning":
            explanation_block = "Leather cleaning removes surface dirt and build up, deep cleans the leather surface. Leather Conditioning moisturizes, softens, strengthens, polishes the leather, and prevents water spotting and cracking. Leather Retouching treats minor scuffs and discoloration, and renews color finish. Leather Protection applies a finish protection."
        elif service_type == "Leather Reupholstery" and item_type in ["Sofa", "Chair", "Cushion"]:
            explanation_block = "Our Full Leather Upholstery service replacing all upholstery with new leather of your choice. We offer a wide selection of colors, textures, and finishes. If only certain cushions or sections are being replaced, our Partial Leather Upholstery service recovers damaged leather with closely matching leather. This requires purchasing new leather and disassembly of the upholstery."
        elif service_type == "Leather Dyeing (Color Change)" and item_type in ["Sofa", "Chair", "Cushion", "Bag", "Coat"]:
            explanation_block = f"This service treats the old finish and dyes the leather in your selected color — {color_selection}. It also refreshes the overall finish of the {item_type}, enhancing both appearance and longevity. We complete the process with a protective topcoat to prevent color transfer."
            if color_selection == "Other":
                explanation_block += " Feel free to browse available color options for your item here: https://www.releather.com/services/leather-dyeing#ColorOptions"
        elif service_type == "Foam Replacement & Restuffing":
            if item_type in ["Sofa", "Chair"]:
                explanation_block = "This service replaces the seat cushion core with high-resilience (HR) grade foam and adds polyester fiber padding for a fuller, more structured look. We offer HR foam in soft, medium, and firm densities to suit your comfort preference."
            elif item_type == "Cushion":
                explanation_block = """This service replaces the seat cushion core with high-resilience (HR) grade foam and adds polyester fiber padding for a fuller, more structured look. We offer HR foam in soft, medium, and firm densities to suit your comfort preference.

Foam Replacement Pricing 
Standard Sofa Seat Cushion: $300–450 each 
Reference dimensions: 
– Thickness: 4" to 6" 
– Width: 22" to 26"
– Depth: 20" to 24"

Larger Seat Cushions: $475–$550 each 
Common for oversized or deep-seat sofas. 
Reference dimensions: 
– Thickness: 5" to 8" 
– Width: 26" to 32" 
– Depth: 24" to 34" """


        # --- Populate Disclaimer Block ---
        if service_type == "Leather Restoration" and item_type == "Cushion":
            disclaimer_block += "Treating a single section or cushion may result in a visible mismatch with the rest of the upholstery. While we strive to achieve the best color match, existing patina can prevent a completely seamless blend. As an alternative, we recommend Full Leather Restoration, which treats the entire piece, addressing all surface wear and restoring a uniform color and finish."
        if service_type == "Leather Reupholstery" and item_type in ["Sofa", "Chair", "Cushion"]:
            disclaimer_block += "We source a leather that closely matches your original, but the worn-in patina of your existing leather may not match the new material seamlessly."
            if is_out_of_state: # This rule concatenates
                disclaimer_block += " To ensure accurate measurements and pattern matching, we require the original seat cover for each unique cushion size mailed to us."
        if service_type == "Leather Dyeing (Color Change)" and item_type in ["Sofa", "Chair", "Cushion", "Bag", "Coat"]:
            # Concatenate if multiple disclaimers apply - ensure proper spacing.
            if disclaimer_block: disclaimer_block += " "
            disclaimer_block += "The new surface coating applied during dyeing may reduce the suppleness of the leather. Accent stitching will be dyed to match the new leather color. While we carefully mask the fabric strip and lining during restoration, some dye transfer may occur. We take precautions to minimize this."

        # --- Populate Estimate Block ---
        if service_type == "Leather Restoration":
            if item_type == "Car":
                estimate_block = "$550–$650 per seat. Discount available for multiple seats or full interior"
            elif item_type == "Sofa":
                estimate_block = "$1800-2200 per sofa."
            elif item_type == "Chair":
                estimate_block = "$800-1200 per chair."
            elif item_type == "Cushion":
                estimate_block = "$450-650 per cushion."
        elif service_type == "Leather Cleaning & Conditioning":
            if item_type == "Sofa":
                estimate_block = "$900-1200 per sofa."
            elif item_type == "Chair":
                estimate_block = "$600-800 per chair."
            elif item_type == "Cushion":
                estimate_block = "$450-650 per cushion."
            elif item_type == "Car":
                estimate_block = "$450-650 per seat."
            elif item_type in ["Bag", "Coat"]:
                estimate_block = "$250-350 per item."
        elif service_type == "Leather Reupholstery":
            if item_type == "Sofa":
                estimate_block = "$6500-8500 per sofa."
            elif item_type == "Chair":
                estimate_block = "$3500-4500 per chair."
            elif item_type == "Cushion":
                if cushions_detachable_fixed == "Detachable":
                    estimate_block = "$850-1200 per cushion."
                elif cushions_detachable_fixed == "Fixed":
                    estimate_block = "$1200-1400 per cushion."
        elif service_type == "Leather Dyeing (Color Change)":
            if item_type == "Sofa":
                estimate_block = "$2500-2800 per Sofa."
            elif item_type == "Chair":
                estimate_block = "$1600-1800 per Chair."
            elif item_type == "Cushion":
                estimate_block = "$550-650 per Cushion." # Corrected from "Chair" to "Cushion"
            elif item_type == "Bag":
                estimate_block = "$350"
            elif item_type == "Coat":
                estimate_block = "$550"
        elif service_type == "Foam Replacement & Restuffing":
            if item_type == "Sofa":
                estimate_block = "$1200-1500 per sofa. Additional labor cost for fixed seating."
            elif item_type == "Chair":
                estimate_block = "$850-950 per chair. Additional labor cost for fixed seating." # Corrected from "sofa" to "chair"
            elif item_type == "Cushion":
                estimate_block = "$350-450 per cushion. Additional labor cost for fixed seating."

        # --- Populate Completion Block ---
        if service_type == "Leather Restoration" and item_type == "Car":
            completion_block = "1-day turnaround. All work is completed at our shop."
        elif service_type == "Leather Restoration" and item_type in ["Sofa", "Chair", "Cushion"]:
            completion_block = "2 weeks."
        elif service_type in ["Leather Cleaning & Conditioning", "Leather Dyeing (Color Change)"] and \
             item_type in ["Sofa", "Chair", "Cushion", "Bag", "Coat"]:
            completion_block = "1-2 weeks."
        elif service_type == "Leather Reupholstery":
            if item_type in ["Sofa", "Chair"]:
                completion_block = "3-4 weeks."
            elif item_type == "Cushion": # This applies if it's reupholstery and a cushion
                completion_block = "2 weeks."
        elif service_type == "Leather Dyeing (Color Change)" and item_type in ["Sofa", "Chair", "Bag", "Coat"]:
            completion_block = "2 weeks."
        elif service_type == "Foam Replacement & Restuffing" and item_type in ["Sofa", "Chair", "Cushion"]:
            completion_block = "2 weeks."

        # --- Populate Delivery Block ---
        if service_type == "Leather Restoration" and item_type == "Car":
            delivery_block = """Service location:
ReLeather
751 S State College Blvd, Unit 38
Fullerton, CA 92831"""
        elif service_type in ["Leather Restoration", "Leather Dyeing (Color Change)", "Leather Cleaning & Conditioning"] and \
             item_type in ["Bag", "Coat"]:
            delivery_block = "Drop-off by appointment at our Fullerton, CA shop. Non-local clients can ship items. Return Shipping is $25 or waived with a prepaid return label."
        elif zip_code is not None and 90000 <= zip_code <= 92899 and item_type in ["Sofa", "Chair"]:
            delivery_block = "Free Pick-up and delivery in Orange County."
        elif county_name: # Check if a specific county was identified
            delivery_block = f"Pick-up and delivery available in {county_name} for $200 (round trip)."
        elif is_out_of_state:
            delivery_block = """Return shipping is quoted separately.

Shipping instructions for mailed-in orders will be provided after confirming your order."""

        # Construct the email subject directly in Python
        email_subject = f"{service_type} Estimate – ReLeather" if service_type else "ReLeather Service Inquiry"

        # --- Construct the simplified Gemini Prompt ---
        # Gemini now only receives the template and pre-filled blocks, without (Block names).
        # The subject is handled entirely by Python. Removed extra <br/> tags as requested.
        gemini_prompt = f"""Use the ReLeather Email Template verbatim to respond to a customer inquiry from Fillout form submission data. Do not paraphrase any of the provided wording. Keep formatting as is and use HTML <br/> tags for line breaks.

Here is the ReLeather Email Template:
  
Hi {customer_first_name}, <br/>
Thank you for your interest in ReLeather.<br/>
Based on the information provided, we recommend our {service_type} for your {plural_item_type}.<br/>
{explanation_block}<br/>
Please note: {disclaimer_block}<br/>
Estimated cost: {estimate_block}.<br/>
Completion time: {completion_block}<br/>
{delivery_block}<br/>
Please feel free to contact us with any questions or to proceed with your order.
---"""

        # Call Gemini API to generate the email based on the prompt
        logging.info("Sending prompt to Gemini API...")
        response = model.generate_content(gemini_prompt)
        generated_email_body = response.text
        logging.info(f"Received raw email body from Gemini: {generated_email_body}")

        # Clean up any potential markdown formatting from Gemini's response (e.g., code block fences)
        generated_email_body = re.sub(r'```[a-zA-Z]*\n|\n```', '', generated_email_body).strip()
        
        # Replace remaining newlines with HTML breaks for consistency, even if Gemini adds them.
        generated_email_body = generated_email_body.replace('\n', '<br/>')

        # Append the signature block at the end of the generated body
        # Added <br/><br/> for spacing between body and signature
        generated_email_body += "<br/><br/>" + OUTLOOK_EMAIL_SIGNATURE

        logging.info("--- Final Generated Email Content ---")
        logging.info(f"Subject: {email_subject}")
        logging.info(f"Body:\n{generated_email_body}")
        logging.info("-----------------------------------")

        # --- Create Outlook Draft ---
        access_token = get_access_token(AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET)
        if access_token:
            draft_created = create_outlook_draft(access_token, OUTLOOK_SENDER_EMAIL, customer_email, email_subject, generated_email_body)
            if draft_created:
                logging.info(f"Outlook draft successfully created for {customer_email}.")
            else:
                logging.error(f"Failed to create Outlook draft for {customer_email}. Check Graph API configuration and permissions.")
        else:
        # Log this error specifically if token couldn't be obtained
            logging.error("Could not obtain access token for Microsoft Graph API. Outlook draft not created.")


        # Return the generated email content as JSON. Fillout can use this in its email integrations.
        return jsonify({
            "data": {
                "response": generated_email_body,
                "subject": email_subject,
                "customer_email": customer_email # Include customer email if Fillout needs it for sending
            },
            "statusCode": 200
        }), 200

    except Exception as e:
        # Log the full traceback for better debugging
        logging.error(f"Error processing webhook: {e}", exc_info=True)
        return jsonify({"error": "Internal server error", "details": str(e)}), 500

@app.route("/", methods=["GET"])
def index():
    return "Webhook server is running."

"""
# --- Flask Webhook Route ---
Paste your webhook logic here exactly as before.
Do NOT modify anything above this line.
"""

# -----------------------------------------------------
#  FLASK APP RUN (LOCAL)
# -----------------------------------------------------

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
