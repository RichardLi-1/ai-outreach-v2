import json
from openai import OpenAI, Timeout
from settings import settings

client = OpenAI(api_key=settings.openai_api_key, timeout=Timeout(120, connect=10))

# Load JSONL and convert to JSON array
with open("data.jsonl", "r") as f:
    records = [json.loads(line) for line in f if line.strip()]

json_bytes = json.dumps(records, indent=2).encode("utf-8")

vs = client.vector_stores.create(name="Alberta GIS Managers")

batch = client.vector_stores.file_batches.upload_and_poll(
    vector_store_id=vs.id,
    files=[("data.json", json_bytes, "application/json")]
)

print(f"Vector store ID: {vs.id}")
print(f"Status: {batch.status}")
print(f"File counts: {batch.file_counts}")
