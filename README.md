# AI Outreach

A Windows desktop application for enriching government contact spreadsheets using AI-powered web search, email finding, and document retrieval.

---

## Overview

AI Outreach processes Excel/CSV files containing lists of government entities (counties, cities, villages) and automatically finds the relevant contact person — GIS Manager, Mayor/County Manager, or Property Assessor — along with their email, phone number, role, and website. It writes results back to a new output file without modifying the original.

A separate Alberta-specific variant (`alberta rag.py`) extends this with a RAG (Retrieval-Augmented Generation) pipeline using a LangChain agent for smarter, document-grounded lookups.

---

## Features

- **Multi-role search** — find GIS Managers, Mayors/County Managers, or Property Assessors per sheet
- **Multi-sheet & stacked-section support** — detects duplicate headers and processes each section independently
- **Automatic column detection** — maps spreadsheet columns to expected fields without manual configuration
- **Email finding via Hunter.io** — uses first name, last name, and government domain to find and verify emails
- **Email confidence scoring** — only overwrites existing emails when Hunter.io returns a higher-confidence result
- **Alternative email preservation** — original email moved to an "Alternative Email" column before overwriting
- **Output file naming** — files named by state/province, role tag, and timestamp (e.g. `AB_NG911_20260310_143022.csv`)
- **Cancellation support** — run can be cancelled mid-way; partial results saved as `_incomplete` files
- **Live log viewer** — collapsible scrollable log panel with real-time output
- **Run statistics** — reports rows processed and files written on completion
- **Dark/light mode** — automatically matches Windows system theme
- **DPI scaling** — correct rendering at non-100% display scaling on Windows
- **Configurable prompts** — GIS, Mayor, and Assessor prompts editable through the Settings panel

### Alberta RAG Variant (`alberta rag.py`)

- **LangChain agent orchestration** — autonomous 3-step pipeline per row using `create_agent`
- **County resolution** — looks up which county/municipal district each village belongs to via web search
- **RAG document lookup** — queries an OpenAI vector store (pre-ingested PDF) for GIS manager contacts
- **Confidence-gated fallback** — only falls back to live web search if RAG confidence < 0.7
- **Structured output** — agent returns a typed `GISContact` Pydantic model ensuring consistent field schema

---

## Tech Stack

| Layer | Technology |
|---|---|
| GUI | Python `tkinter` + [Sun Valley theme](https://github.com/rdbende/Sun-Valley-ttk-theme) (`sv_ttk`) |
| Data processing | `pandas` |
| AI search | OpenAI `gpt-4o-mini-search-preview` (web search built-in) |
| Email finding | [Hunter.io](https://hunter.io) Email Finder + Verifier API |
| RAG (Alberta) | OpenAI Responses API + File Search (vector store) |
| Agent (Alberta) | LangChain `create_agent` (LangGraph-based, v1.x) |
| Config | `pydantic-settings` with `.env` file |
| Theming | `darkdetect`, `pywinstyles` |
| Packaging | PyInstaller |

---

## Setup

### Requirements

- Python 3.11+
- Windows (uses `winsound`, Windows DPI APIs)

### Install dependencies

```bash
pip install -r requirements.txt
```

### Configure `.env`

Create a `.env` file in the project root:

```env
OPENAI_API_KEY=sk-...
HUNTER_API_KEY=...
MAX_TOKENS=1000

INITIAL_PROMPT=...
INITIAL_PROMPT_MAYOR=...
INITIAL_PROMPT_ASSESSOR=...
PROMPT_FORMAT_GIS=...
PROMPT_FORMAT_ASSESSOR=...

# Alberta RAG only
PROMPT_FIND_COUNTY=...
PROMPT_FIND_IN_FILE=...
FILE_ID=vs_...
```

### Run

```bash
# Standard version
python main.py

# Alberta RAG version
python "alberta rag.py"
```

---

## Usage

1. Click **Select File** and choose an Excel (`.xlsx`) or CSV (`.csv`) file
2. Select an **output folder**
3. For each sheet/section, confirm column mappings and choose a role to search for
4. The app processes each row, writes enriched results, and plays a sound on completion
5. Output files appear in the selected folder named by state, tag, and timestamp

---

## Output Columns Written

| Column | Source |
|---|---|
| First Name, Last Name | OpenAI search |
| Role/Title | OpenAI search |
| Email | OpenAI search or Hunter.io (higher confidence wins) |
| Phone Number | OpenAI search or Hunter.io |
| Email Domain | OpenAI search |
| Email Confidence | Hunter.io score |
| Alternative Email | Previous email (if overwritten) |
| Hunter Email Source | Hunter.io source URL |
| Source | Government website found by OpenAI |
| Contact Tag | Role tag (NG911, QQ) |

---

## Project Structure

```
main.py               # Main application (US/general)
alberta rag.py        # Alberta variant with LangChain RAG agent
alberta_tools.py      # LangChain tool definitions (county lookup, RAG, web search)
openai_hunter_client.py  # OpenAI search + Hunter.io API wrappers
presets.py            # Role enum, state/province correction map
settings.py           # Pydantic settings loaded from .env
```
