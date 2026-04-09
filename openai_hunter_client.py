from settings import settings
from openai import OpenAI, Timeout
import requests
from presets import Role, SearchFor
import logging
import re
import json

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

def search_misc(prompt: str, searchType: SearchFor, name="", role="", company="") -> str: #for searching population, etc
    logger.info(f"Calling OpenAI to search for {prompt}")

    match searchType:
        case SearchFor.POPULATION:
            message=[{"role": "system", "content": settings.prompt_find_population}]
        case SearchFor.OUTREACH_MESSAGE:
            message=[{"role": "system", "content": settings.prompt_find_outreach_message}]
        case SearchFor.HAS_GIS_DEPARTMENT:
            message=[{"role": "system", "content": settings.prompt_has_gis_department}]
        

    message.append({"role": "user", "content": prompt})

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


def ingest_documents(file_paths, store_name):
    files = [open(path, "rb") for path in file_paths]
    vs = client.vector_stores.create(name=store_name)

    try:
        batch = client.vector_stores.file_batches.upload_and_poll(
            vector_store_id=vs.id,
            files=files
        )
    except Exception as e:
        logger.error(f"Error during file ingestion: {e}")

    finally:
        for f in files:
            f.close()

    print(f"Vector store ID: {vs.id}")
    print(f"Status: {batch.status}")
    print(f"File counts: {batch.file_counts}")
    return vs.id

def query_rag(vs_id: str, county: str, state: str) -> str:
    """Search the document. Returns a JSON string with firstName, lastName, email, phoneNumber, role, govWebsite, and confidence (0.0-1.0). Use this before falling back to web search."""
    response = client.responses.create(
        model="gpt-4o-mini",
        instructions=settings.prompt_find_in_file,
        input=f"{county}, {state}",
        tools=[{
            "type": "file_search",
            "vector_store_ids": [vs_id]
        }]
    )
    for item in response.output:
        if item.type == "file_search_call":
            logger.info(f"[RAG] file_search status: {item.status}")
            if hasattr(item, 'results') and item.results:
                for r in item.results:
                    logger.info(f"[RAG] hit score={r.score:.3f} | {str(r.text)[:200]}")
            else:
                logger.warning("[RAG] file_search returned no results")
    result = response.output_text
    logger.info(f"[RAG] model output: {result}")
    return result


"""



if __name__ == "__main__":
    #print("Testing lookup county")
    #content = _lookup_county("Vegreville", "AB - Alberta", settings.prompt_find_county)
    #print(content)

    #print("Testing file ingestion")
    #ingest_file("alberta.pdf")

    print("Testing assistant creation")
    create_assistant("vs_69b074dcd92081918de4d72158657bce")"""