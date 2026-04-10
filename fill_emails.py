"""
fill_emails.py
--------------
Fills missing emails in the merged contact list by:
1. Copying any known emails from 'leap2026 contact list.xlsx' (matched by first+last name)
2. Calling Hunter.io (via openai_hunter_client) for still-missing emails
3. Using OpenAI to resolve a domain when only an org name is available
Writes results to a new timestamped xlsx in the same folder.
"""

import re
import sys
import io
import time
import pandas as pd
import hunter_client

# Fix Windows console encoding so Unicode chars don't crash
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf-8-sig"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from pathlib import Path
from datetime import datetime

# Reuse existing clients
import openai_client

MERGED_FILE = Path(r"C:\dev\ai-outreach\merged_split_hunter_leap2026 contact list_20260316_155137_20260316_160246_20260316_203518.xlsx")
LEAP_FILE   = Path(r"C:\dev\ai-outreach\leap2026 contact list.xlsx")

# ── helpers ──────────────────────────────────────────────────────────────────

def normalize_name(s):
    """Lowercase, strip whitespace/non-breaking spaces for matching."""
    return re.sub(r"\s+", " ", str(s).replace("\xa0", " ").replace("\u200b", "").strip()).lower()

def extract_domain(val):
    if not val or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    s = re.sub(r"^https?://", "", s)
    s = re.sub(r"^www\.", "", s)
    s = s.split("/")[0].split("?")[0].split("#")[0]
    return s if "." in s and " " not in s.strip() else None

def extract_domain_from_text(text):
    """Extract a bare domain from verbose OpenAI prose (e.g. '...is wake.gov...')."""
    if not text:
        return None
    # Try to find a domain pattern like word.tld or word.word.tld
    matches = re.findall(r'\b([a-z0-9](?:[a-z0-9\-]{0,61}[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9\-]{0,61}[a-z0-9])?)+)\b', text.lower())
    # Filter out things that aren't real domains (must have a proper TLD)
    tlds = {
        "gov", "com", "org", "net", "edu", "ca", "uk", "au", "us", "info",
        "io", "co", "mil", "int", "biz", "nz", "de", "fr", "jp", "ch",
    }
    for m in matches:
        parts = m.split(".")
        if len(parts) >= 2 and parts[-1] in tlds:
            return m
    return None

def is_blank(val):
    return val is None or (isinstance(val, float) and pd.isna(val)) or str(val).strip() in ("", "nan", "NaN")

# ── load files ────────────────────────────────────────────────────────────────

print("Loading files…")
df = pd.read_excel(MERGED_FILE)
leap = pd.read_excel(LEAP_FILE)

# Ensure output columns exist
for col in ("Email", "Email Confidence", "Email Domain", "Company Website", "Hunter Email Source"):
    if col not in df.columns:
        df[col] = ""

# ── step 1: pull emails from leap2026 list ────────────────────────────────────

leap_email_map = {}  # normalized "first last" → email
name_col = "Name\xa0"
for _, row in leap.iterrows():
    raw = str(row.get(name_col, "") or "")
    email = row.get("Email", "")
    if not is_blank(email):
        key = normalize_name(raw)
        leap_email_map[key] = str(email).strip()

print(f"Leap2026 emails available: {len(leap_email_map)}")

filled_from_leap = 0
for idx, row in df.iterrows():
    if not is_blank(row.get("Email")):
        continue
    first = str(row.get("First Name", "") or "").strip()
    last  = str(row.get("Last Name", "") or "").strip()
    key = normalize_name(f"{first} {last}")
    if key in leap_email_map:
        df.at[idx, "Email"] = leap_email_map[key]
        print(f"  Leap fill: {first} {last} → {leap_email_map[key]}")
        filled_from_leap += 1

print(f"Filled {filled_from_leap} emails from leap2026 list.")

# ── step 2: Hunter.io for remaining gaps ──────────────────────────────────────

still_missing = df[df["Email"].apply(is_blank)].index.tolist()
print(f"\n{len(still_missing)} rows still need emails. Starting Hunter.io lookups…\n")

filled_from_hunter = 0
skipped = 0

for idx in still_missing:
    row = df.loc[idx]
    first = str(row.get("First Name", "") or "").replace("\xa0", " ").replace("\u200b", "").strip()
    last  = str(row.get("Last Name", "") or "").replace("\xa0", " ").replace("\u200b", "").strip()

    # Some rows have the full name fused with \xa0 in First Name; try to split them
    if (is_blank(last) or last.lower() == "nan") and " " in first:
        parts = first.split(" ", 1)
        first, last = parts[0].strip(), parts[1].strip()

    # Skip rows with broken/missing names
    if is_blank(first) or is_blank(last) or last.lower() == "nan":
        print(f"Row {idx}: skipping — missing first or last name ({first!r} {last!r})")
        skipped += 1
        continue

    # Resolve domain: prefer Email Domain > Company Website > Organization
    domain = (
        extract_domain(row.get("Email Domain"))
        or extract_domain(row.get("Company Website"))
        or extract_domain(row.get("Organization"))
    )

    org_raw = str(row.get("Organization", "") or "").strip()

    if not domain:
        if is_blank(org_raw):
            print(f"Row {idx}: skipping — no domain or organization for {first} {last}")
            skipped += 1
            continue
        print(f"Row {idx}: asking OpenAI for domain of '{org_raw}'…")
        raw_domain = openai_client.find_domain(org_raw)
        # find_domain may return verbose prose; extract clean domain
        domain = extract_domain(raw_domain) or extract_domain_from_text(raw_domain or "")
        if not domain:
            print(f"Row {idx}: could not resolve domain — skipping {first} {last}")
            skipped += 1
            continue
        print(f"Row {idx}: OpenAI resolved domain → {domain}")

    print(f"Row {idx}: Hunter lookup — {first} {last} @ {domain}")

    try:
        res = hunter_client.find_email(first, last, domain)
        attempts = 1
        while str(res[0]) == "202" and attempts <= 5:
            time.sleep(2)
            res = hunter_client.find_email(first, last, domain)
            attempts += 1

        if str(res[0]) == "200":
            data   = res[1].get("data") or {}
            email  = data.get("email") or ""
            score  = data.get("score", "")
            sources = data.get("sources") or []
            source_uri   = sources[0]["uri"] if sources else ""
            email_domain = data.get("domain") or domain

            df.at[idx, "Email"]               = email
            df.at[idx, "Email Confidence"]    = score if score != "" else ""
            df.at[idx, "Email Domain"]        = email_domain
            df.at[idx, "Company Website"]     = org_raw
            df.at[idx, "Hunter Email Source"] = source_uri

            print(f"  → {email or '(no email found)'} (confidence: {score})")
            if email:
                filled_from_hunter += 1
        else:
            print(f"  → Hunter.io returned status {res[0]}")
    except Exception as e:
        print(f"Row {idx}: ERROR — {e}")

# ── save ──────────────────────────────────────────────────────────────────────

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
out_path = MERGED_FILE.parent / f"filled_{MERGED_FILE.stem}_{timestamp}.xlsx"
df.to_excel(out_path, index=False, engine="openpyxl")

print(f"\n{'─'*60}")
print(f"Done.")
print(f"  From leap2026 list : {filled_from_leap}")
print(f"  From Hunter.io     : {filled_from_hunter}")
print(f"  Skipped            : {skipped}")
print(f"  Output             : {out_path}")
