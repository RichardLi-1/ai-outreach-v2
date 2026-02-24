from settings import settings
from openai import OpenAI, Timeout
import requests
from presets import Role
import logging

logger = logging.getLogger(__name__)
client = OpenAI(api_key=settings.openai_api_key,
                timeout=Timeout(60, connect=10),
    max_retries=4,)

def search(prompt: str, role: Role, system_prompt: str) -> str:
    logger.info(f"Calling OpenAI to search for {role} in {prompt}")
    message=[]

    message=[{"role": "system", 
                    "content": system_prompt}]
    
    message.append({
        "role": "user",
        "content": prompt
    })

    chat = client.chat.completions.create(
        model="gpt-4o-mini-search-preview",
        messages=message,
        max_tokens = settings.max_tokens,
    )
    
    if not chat.choices:
        logger.error("Empty response from OpenAI")
        return None
    logger.info(chat.choices[0].message.content)
    return chat.choices[0].message.content

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