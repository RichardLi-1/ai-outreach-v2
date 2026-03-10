from langchain_core.tools import tool
import openai_hunter_client
from settings import settings
import logging
from openai import OpenAI, Timeout

logger = logging.getLogger(__name__)

VECTOR_STORE_ID = settings.file_id

#Alberta

client = OpenAI(api_key=settings.openai_api_key, timeout=Timeout(60, connect=10), max_retries=4)


@tool
def lookup_county(town: str, state: str) -> str:
    """Find which county or municipal district a town is located in for a given Canadian province. Returns only the county name."""
    logger.info(f"Calling OpenAI to search for county of {town} in {state}")
    message=[]
    prompt = settings.initial_prompt + ": " + town + ", " + state

    message.append({
        "role": "user",
        "content": prompt
    })

    chat = client.chat.completions.create(
        model="gpt-4o-mini-search-preview",
        messages=message,
        max_tokens = 100,
    )
    
    print(chat)
    if not chat.choices:
        logger.error("Empty response from OpenAI")
        return None
    logger.info(chat.choices[0].message.content)
    return chat.choices[0].message.content

@tool
def query_rag(county: str) -> str:
    """Search the internal document database for GIS manager contact info for an Alberta county. Returns a JSON string with firstName, lastName, email, phoneNumber, role, govWebsite, and confidence (0.0-1.0). Use this before falling back to web search."""
    response = client.responses.create(
        model="gpt-4o-mini",
        instructions=settings.prompt_find_in_file,
        input=f"GIS manager for {county}, Alberta",
        tools=[{
            "type": "file_search",
            "vector_store_ids": [settings.file_id]
        }]
    )
    result = response.output_text  # then json.loads()
    return result



@tool
def web_search_gis(county: str) -> str:
    """Search the web for the GIS manager contact info for an Alberta county government. Use this only if query_rag returns a confidence below 0.7 or finds nothing. Returns a JSON string with firstName, lastName, email, phoneNumber, role, and govWebsite."""
    logger.info(f"Calling OpenAI to search for GIS Manager in {county}, Alberta")
    message=[]

    message=[{
        "role": "system",
        "content": settings.initial_prompt + settings.prompt_format_gis + " If none found, return exactly the string None and nothing else. Do not explain or elaborate."
    }]


    message.append({
        "role": "user",
        "content": county
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



"""def ingest_file(file_path: str) -> str:
    vs = client.vector_stores.create(name="Alberta GIS Managers")

    client.vector_stores.file_batches.upload_and_poll(
        vector_store_id=vs.id,
        files=[open(file_path, "rb")]
    )
    print("id:" + vs.id)
    print("File ingested and vectorized successfully.")

    #Deprecated

def create_assistant(vs_id: str) -> str:
    assistant = client.beta.assistants.create(
        name="Alberta GIS RAG",
        model="gpt-4o-mini",
        tools=[{"type": "file_search"}],
        tool_resources={"file_search": {"vector_store_ids": [vs_id]}},
        instructions="You are a lookup tool for Alberta GIS manager contacts. When asked about a county, search the document and return a JSON object with fields: firstName, lastName, email, phoneNumber, role, govWebsite, confidence (0.0-1.0). If not found, return {\"confidence\": 0}."
    )
    # save assistant.id
    print("Assistant created with ID: " + assistant.id)
    return assistant.id"""
