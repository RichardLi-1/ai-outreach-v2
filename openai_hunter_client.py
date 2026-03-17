from settings import settings
from openai import OpenAI, Timeout
import requests
from presets import Role
import logging
import re

logger = logging.getLogger(__name__)
client = OpenAI(api_key=settings.openai_api_key,
                timeout=Timeout(60, connect=10),
    max_retries=4,)

def search(prompt: str, role: Role, system_prompt: str) -> str:
    logger.info(f"Calling OpenAI to search for {role} in {prompt}")
    message=[]

    if role == Role.ASSESSOR:
        message=[{"role": "system", 
                    "content": system_prompt + settings.prompt_format_assessor}]
    else:
        message=[{"role": "system", 
                    "content": system_prompt + settings.prompt_format_gis}]


    message.append({
        "role": "user",
        "content": prompt
    })

    chat = client.chat.completions.create(
        model="gpt-4o-mini-search-preview",
        messages=message,
        max_tokens = settings.max_tokens,
    )
    
    print(chat)
    if not chat.choices:
        logger.error("Empty response from OpenAI")
        return None
    logger.info(chat.choices[0].message.content)
    return chat.choices[0].message.content

def find_domain(org_name: str):
    """Ask OpenAI to find the official website domain for an organization by name."""
    logger.info(f"Asking OpenAI for domain of: {org_name}")
    try:
        chat = client.chat.completions.create(
            model="gpt-4o-mini-search-preview",
            messages=[
                {"role": "system", "content": "You find the official website domain for organizations. Reply with only the bare domain (e.g. 'example.com'), no protocol, www, or path. If you cannot find it, reply exactly: unknown"},
                {"role": "user", "content": f"What is the official website domain for this organization: {org_name}"}
            ],
            max_tokens=50,
        )
        if not chat.choices:
            return None
        result = chat.choices[0].message.content.strip().lower()
        result = re.sub(r'^https?://', '', result)
        result = re.sub(r'^www\.', '', result)
        result = result.split('/')[0].split('?')[0].strip()
        if '.' in result and result != 'unknown':
            logger.info(f"OpenAI found domain: {result}")
            return result
        return None
    except Exception as e:
        logger.warning(f"find_domain failed for '{org_name}': {e}")
        return None


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


"""



if __name__ == "__main__":
    #print("Testing lookup county")
    #content = _lookup_county("Vegreville", "AB - Alberta", settings.prompt_find_county)
    #print(content)

    #print("Testing file ingestion")
    #ingest_file("alberta.pdf")

    print("Testing assistant creation")
    create_assistant("vs_69b074dcd92081918de4d72158657bce")"""