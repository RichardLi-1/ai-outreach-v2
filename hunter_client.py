import logging
import requests
logger = logging.getLogger(__name__)
import settings


def find_email(firstName, lastName, domain):
    logger.info(f"Called Hunter.io Email Finding API for: {firstName} {lastName}, {domain}")

    findURL = f"https://api.hunter.io/v2/email-finder?domain={domain}&first_name={firstName}&last_name={lastName}&api_key={settings.hunter_api_key}"
    response = requests.get(findURL, timeout=40)
    logger.info(response.status_code)
    logger.info(response.json())

    return (response.status_code, response.json())

def verify_email(email):
    logger.info(f"Called Hunter.io Email Verification API for: {email}")

    validateURL = "https://api.hunter.io/v2/email-verifier?email=" + email + "&api_key=" + settings.hunter_api_key
    response = requests.get(validateURL, timeout=30)
    logger.info(response.status_code)
    logger.info(response.json())

    return (response.status_code, response.json())