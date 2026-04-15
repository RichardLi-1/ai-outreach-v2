# AI Outreach

A Windows desktop application for enriching government contact spreadsheets using AI-powered web search, email verification, and optional document-based lookup.

---

## Overview

AI Outreach processes Excel/CSV files containing government entities (counties, cities, villages) and automatically finds relevant contacts — **GIS Managers** or **Property Assessors** — along with their email, phone, role, and department information. Results are written to a new timestamped output file without modifying the original.

An Alberta-specific variant (`alberta rag.py`) uses a LangChain agent with RAG (Retrieval-Augmented Generation) for smarter, document-grounded contact discovery with fallback to web search.

---

## To build:
Run command:
pyinstaller --onefile --noconsole main.py --add-data ".env;."
on root directory.

## Features

### Core Features
- **GIS Manager & Property Assessor search** — select role per sheet via GUI
- **Multi-sheet & stacked-section support** — auto-detects duplicate headers, processes each section independently
- **Automatic column detection** — recognizes 20+ column aliases (County, City, Email, Phone, First/Last Name, Role/Title, etc.)
- **Email finding & verification** — Hunter.io API for personal email discovery and confidence scoring
- **Confidence-based email overwriting** — only replaces existing emails if Hunter.io returns higher confidence
- **Alternative email preservation** — original email moves to "Alternative Email" column before overwriting
- **Address Department & Website** — extracts GIS department and organization website (for GIS Manager role)
- **Optional population lookup** — toggle to enrich rows with city/county population data via web search
- **Optional outreach messages** — toggle to generate LinkedIn-friendly outreach messages for each contact
- **Output naming** — files auto-named by state, role tag, and timestamp (e.g., `CA_NG911_20260415_143022.csv`)
- **Live log viewer** — collapsible, scrollable panel with real-time processing output
- **Cancellation support** — partial results saved with `_incomplete` suffix if run is cancelled
- **Run statistics** — summary of rows processed and output files created
- **Dark/light theme** — automatically matches Windows system theme
- **DPI scaling** — correct rendering at non-standard display scaling
- **Editable prompts** — customize search prompts via Settings panel

### Alberta RAG Variant (`alberta rag.py`)
- **LangChain agent orchestration** — autonomous 3-step pipeline per row using `create_agent`
- **County resolution** — web search to determine county/municipal district for each location
- **RAG document lookup** — queries OpenAI vector store (pre-ingested PDF) for GIS manager contacts
- **Confidence-gated fallback** — falls back to live web search only if RAG confidence < 0.7
- **Structured output** — returns typed `GISContact` Pydantic model for consistent schema

---

## Tech Stack

| Layer | Technology |
|---|---|
| GUI | Python `tkinter` + [Sun Valley theme](https://github.com/rdbende/Sun-Valley-ttk-theme) (`sv_ttk`) |
| Data processing | `pandas` (3.0.0+) |
| AI search | OpenAI `gpt-4o-mini-search-preview` (web search built-in) |
| Email finding | [Hunter.io](https://hunter.io) API (find & verify) |
| RAG (optional) | OpenAI Responses API + File Search |
| Agent (Alberta) | LangChain 1.x with `create_agent` |
| Config | `pydantic-settings` with `.env` file |
| Theming | `darkdetect`, `pywinstyles` |
| Utilities | `openpyxl`, `tenacity` (retry logic) |
| Packaging | PyInstaller |
| Python | 3.11+ (Windows only)

---

## Setup

### Requirements

- Python 3.11+
- Windows (uses tkinter, Windows theming APIs, DPI scaling)

### Install dependencies

```bash
pip install -r requirements.txt
```

### Configure `.env`

Create a `.env` file in the project root with at minimum:

```env
OPENAI_API_KEY=sk-...
HUNTER_API_KEY=...
MAX_TOKENS=1000

# System prompts (required)
INITIAL_PROMPT=<GIS Manager search prompt>
INITIAL_PROMPT_ASSESSOR=<Property Assessor search prompt>
PROMPT_FIND_DOMAIN=<Domain extraction prompt>
PROMPT_FIND_POPULATION=<Population lookup prompt>
PROMPT_FIND_OUTREACH_MESSAGE=<LinkedIn outreach message prompt>
PROMPT_FIND_HAS_GIS=<GIS department detection prompt>

# Optional: For RAG-based lookups only
FILE_ID=vs_...  # Vector store ID from OpenAI file_search
```

For the **Alberta RAG variant** (`alberta rag.py`), you may also set:
```env
PROMPT_FIND_COUNTY=<County resolution prompt>
PROMPT_FIND_IN_FILE=<RAG document query prompt>
```

### Run

```bash
# Main application (GIS Manager & Property Assessor search)
python main.py

# Alberta RAG variant (LangChain agent-based)
python "alberta rag.py"
```

---

## Usage

### Main Application (`main.py`)

1. Click **Select File** and choose an Excel (`.xlsx`) or CSV (`.csv`) file
2. Select an **output folder** for results
3. Configure optional toggles in **Settings** (if needed):
   - **Search Population** — enrich rows with city/county population via web search
   - **Find Outreach Messages** — generate LinkedIn-friendly outreach suggestions
4. For each sheet/section detected:
   - Confirm column mappings (auto-detected)
   - Select a **Role**: GIS Manager or Property Assessor
   - Enter a **Role Tag** (e.g., "NG911", "QQ") — used in output filename
5. Processing begins automatically:
   - Each row is enriched with contact info via OpenAI web search
   - Hunter.io verifies and finds personal emails
   - Optional: population and outreach messages are added
   - Partial results are shown in the log viewer
6. A sound plays on completion; output files appear in your output folder
7. If cancelled mid-run, partial results are saved as `_incomplete` files

### Alberta RAG Application (`alberta rag.py`)

Runs the same enrichment process but uses a **LangChain agent** that:
1. Resolves the county/district for each location via web search
2. Queries your ingested PDF vector store for GIS manager contact info
3. Falls back to web search if vector store confidence < 0.7
4. Returns structured contact data

**Note:** Requires `FILE_ID` (vector store ID) in `.env` to be useful.

---

## Output Columns

The following columns are written to the output file:

| Column | Source | Notes |
|---|---|---|
| First Name | OpenAI web search | Parsed from contact name |
| Last Name | OpenAI web search | Parsed from contact name |
| Email | OpenAI or Hunter.io | Hunter email wins if higher confidence |
| Phone Number | OpenAI search | From government website |
| Role/Title | OpenAI web search | e.g., "GIS Manager", "Property Assessor" |
| Email Domain | OpenAI search | Extracted from email; filters generic domains (gmail, yahoo, etc.) |
| Email Confidence | Hunter.io | Confidence score from email verification (0–100) |
| Alternative Email | Preserved from input | Original email if overwritten by Hunter.io |
| Organization Website | OpenAI search | Government entity's main website |
| Address/Department | OpenAI search | For GIS Manager role only; e.g., "GIS Department" |
| Hunter Email Source | Hunter.io | Source URL of email discovery |
| Source | OpenAI search | Government website where contact was found |
| Contact Tag | User-selected | Role tag from GUI (e.g., "NG911", "QQ") |
| Population | OpenAI search | *Optional*; only if "Search Population" toggle enabled |
| Contact LinkedIn Outreach Message | OpenAI search | *Optional*; only if "Find Outreach Messages" toggle enabled |

---

## Project Structure

```
main.py                    # Main GUI application (tkinter)
alberta rag.py             # Alberta variant with LangChain RAG agent
alberta_tools.py           # LangChain tool definitions (county lookup, RAG, web search)
openai_client.py           # OpenAI API client (web search, RAG, misc queries)
hunter_client.py           # Hunter.io API client (email finding & verification)
utilities.py               # Spreadsheet processing (column detection, splitting)
presets.py                 # Role enums, state/province name corrections
settings.py                # Pydantic settings (loads from .env)
app_state.json             # GUI state persistence (window size, last paths)
requirements.txt           # Python dependencies
.env                       # Configuration (API keys, prompts) — not in repo
```

### Utility Tools (Accessible from GUI or CLI)

| Tool | File | Purpose |
|---|---|---|
| **Hunter Finder** | `hunter_finder.py` | Batch email lookup via Hunter.io API; useful for pre-processing contacts |
| **Name Splitter** | `name_splitter.py` | Split "Full Name" into "First Name" and "Last Name" columns |
| **Merge Tool** | `merge.py` | Combine two CSV/Excel files with duplicate removal |
| **Hunter CLI** | `hunter_only.py` | Interactive command-line Hunter.io email lookup |
| **Email Filler** | `fill_emails.py` | Fill missing emails by looking up from a reference file |
| **RAG Ingest** | `ingest.py` | Convert JSONL to JSON format and upload to OpenAI vector store |

---

## Implementation Details

### Column Detection

The app recognizes 20+ column name aliases and auto-maps them:

**Location** — County, City, Municipality, District, Village, Town, etc.  
**Names** — First Name, Last Name, Full Name, etc.  
**Contact** — Email, Phone, Phone Number, Website, Organization Website, etc.  
**Optional** — Population, LinkedIn, LinkedIn Profile, Alternative Email, etc.

If a column is ambiguous or missing, the GUI prompts you to confirm the mapping before processing.

### Duplicate Header Detection

When a spreadsheet contains stacked datasets (multiple sections with repeated headers), the app automatically:
1. Detects each header row
2. Splits into independent sections
3. Processes each section separately with independent role selection
4. Writes each to its own output file

### Email Confidence Logic

Hunter.io returns a confidence score (0–100) for each email found:
- If **no email exists** in the input row, Hunter email is always used
- If **email exists** and Hunter confidence > existing confidence, Hunter email overwrites it
- The previous email is preserved in the "Alternative Email" column
- Email confidence score is always written to output

### Hunter.io Retry Logic

Hunter.io email finding is asynchronous and may return HTTP 202 ("pending"):
- The client automatically retries up to 5 times with exponential backoff
- If still pending after 5 attempts, the row is skipped for that email lookup

### Optional Features (via Settings)

**Population Lookup** — If enabled, calls OpenAI to find city/county population for each row  
**Outreach Messages** — If enabled, generates a LinkedIn-friendly outreach message for each contact  
**GIS Department Detection** — Determines if a contact has a dedicated GIS department (prompts are customizable)

### State Persistence

The GUI saves state to `app_state.json`:
- Window size and position
- Last selected file/output folder
- Scroll position in log viewer
- Other UI preferences

This allows the app to restore your previous session on restart.

### Multi-file Output

One output file is created **per sheet/section/role combination**:
- Format: `{STATE_ABBREV}_{ROLE_TAG}_{TIMESTAMP}.csv` (or `.xlsx`)
- Example: `CA_NG911_20260415_143022.csv`
- If cancelled mid-run, incomplete files are saved as `_incomplete` suffix

---

## Known Limitations & Notes

- **Mayor role** — Enum defined but not exposed in main GUI; use Alberta RAG variant if needed
- **Address Department** — For GIS Manager only; may sometimes return school/hospital domains (accuracy issue)
- **Email Domain** — Generic domains (gmail, yahoo, outlook, hotmail) are filtered out; falls back to organization website if unavailable
- **RAG vector store** — Optional; only improves accuracy if documents are relevant and properly ingested
- **Population lookup** — Adds 1–2 seconds per row; disable if you don't need it
- **Windows only** — Uses tkinter, `darkdetect`, and Windows-specific theming APIs

---

## Support & Troubleshooting

If the app crashes or behaves unexpectedly, check:
1. `.env` file has all required keys (`OPENAI_API_KEY`, `HUNTER_API_KEY`)
2. Your OpenAI account has sufficient credits
3. Hunter.io API quota is not exhausted
4. Input file is valid Excel/CSV with expected headers
5. Windows display scaling is set to 100% or use the DPI settings panel in Settings

For more information, see [FEATURES_TO_ADD.md](FEATURES_TO_ADD.md) for planned enhancements.

