# utilities.py
import pandas as pd

def _detect_columns(df):
    cols = {}
    col_list = []
    for column in df.columns:
        # Convert to string in case column name is numeric
        column_str = str(column) if not isinstance(column, str) else column

        # Normalize by removing whitespace and converting to lowercase
        normalized = ''.join(column_str.lower().split())

        # Match columns
        _ALIASES = {
            "County/City": ["county", "county/city", "city/county"],
            "Email": ["email", "contactemail"],
            "Phone Number": ["number", "phonenumber", "contactphonenumber", "contactnumber", "contactphone"],
            "First Name": ["firstname", "contactfirstname", "first"],
            "Last Name": ["lastname", "contactlastname", "last", "surname"],
            "Role/Title": ["position", "role", "title", "role/title", "title/role", "jobtitle"],
            "LinkedIn":     ["linkedin", "contactlinkedin", "linkedinprofile", "contactlinkedinprofile"],
            "Outreach Message": ["contactlinkedinoutreachmessage", "linkedinoutreachmessage", "outreachmessage"],
            "Email Domain": ["emaildomain"],
            "Has GIS Department": ["hasgisdepartment"]
        }

        _COLUMN_MAP = {alias: canonical
                    for canonical, aliases in _ALIASES.items()
                    for alias in aliases
        }

        if "population" in normalized:
            cols["Population"] = column
        elif "contact" in normalized and ("state" in normalized or "province" in normalized):
            cols["Contact State"] = column
        elif normalized in ["state", "province", "state/province", "province/state", "provinceorstate"] or "state" in normalized or "province" in normalized:
            if "State" not in cols:  # Don't overwrite if we already have a state column from an exact match:
                cols["State"] = column
        elif normalized in _COLUMN_MAP:
            cols[_COLUMN_MAP[normalized]] = column
        elif normalized in ["contactlinkedinprofile", "contactlinkedin", "linkedin", "linkedinprofile"] or "linkedin" in normalized:
            cols["LinkedIn"] = column
        elif "contact" in normalized and "tag" in normalized:
            cols["Contact Tag"] = column
        elif "tag" in normalized:
            cols["Tag"] = column
        
        if column is not None and str(column).strip() != "":
            col_list.append(column)
    return cols, col_list

def _find_duplicate_headers(df):
    """
    Find rows that look like header rows (indicating stacked datasets).
    Returns a list of row indices where headers are found.
    """
    header_rows = []

    # Keywords commonly found in headers (single words only for whole-word matching)
    header_keywords = ['country', 'county', 'city', 'email', 'state', 'province',
                        'contact', 'population', 'phone', 'role', 'title', 'name',
                        'first', 'last', 'position', 'address']

    for idx in range(len(df)):
        row = df.iloc[idx]
        # Convert all values in the row to lowercase strings
        row_values = [str(val).lower().strip() for val in row if pd.notna(val) and str(val).strip()]

        # Count how many header keywords appear in this row
        # A cell is a header cell if MOST of its words are keywords
        keyword_matches = 0
        for val in row_values:
            # Split into words
            val_words = val.replace('/', ' ').replace('-', ' ').split()
            if len(val_words) == 0:
                continue

            # Count how many words in this cell are keywords
            keyword_count_in_cell = sum(1 for word in val_words if word in header_keywords)

            # Cell is a header cell if MAJORITY of its words are keywords (>50%)
            # (e.g., "county" = 1/1 ✓, "county/city" = 2/2 ✓, "brevard county" = 1/2 = 50% ✗)
            if keyword_count_in_cell > len(val_words) * 0.5:
                keyword_matches += 1

        # If 3+ cells contain header keywords, it's likely a header row
        if keyword_matches >= 3:
            header_rows.append(idx)

    return header_rows


def _split_by_duplicate_headers(df, sheet_name, logger):
    """
    Split a dataframe into multiple sections if duplicate headers are found.
    Returns a list of tuples: (section_name, section_dataframe)
    """
    header_indices = _find_duplicate_headers(df)

    # If no headers detected, fall back to treating row 0 as the header
    if len(header_indices) == 0:
        section_df = df.copy()
        section_df.columns = [str(col) if pd.notna(col) else "" for col in section_df.iloc[0]]
        section_df = section_df.iloc[1:].reset_index(drop=True)
        return [(sheet_name, section_df)]
    # If exactly one header found, normalize into a single section
    if len(header_indices) == 1:
        start_idx = header_indices[0]
        section_df = df.iloc[start_idx:].copy()
        section_df.columns = [str(col) if pd.notna(col) else "" for col in section_df.iloc[0]]
        section_df = section_df.iloc[1:].reset_index(drop=True)
        if len(section_df) < 1:
            return []
        return [(sheet_name, section_df)]

    # Log that we found multiple headers
    logger.info(f"**Found {len(header_indices)} sections in sheet '{sheet_name}'**")

    # Split into sections
    sections = []
    for i in range(len(header_indices)):
        start_idx = header_indices[i]
        # End is either the next header or the end of the dataframe
        end_idx = header_indices[i + 1] if i + 1 < len(header_indices) else len(df)

        # Extract this section
        section_df = df.iloc[start_idx:end_idx].copy()

        # Skip if section is too small (less than 2 rows including header)
        if len(section_df) < 2:
            continue

        # Set the first row as column headers and remove it from data
        # Convert all header values to strings to avoid numeric column names
        section_df.columns = [str(col) if pd.notna(col) else "" for col in section_df.iloc[0]]
        section_df = section_df.iloc[1:].reset_index(drop=True)

        # After removing header, check if we still have data
        if len(section_df) < 1:
            continue

        # Create a name for this section
        section_name = f"{sheet_name}_part{i + 1}"
        sections.append((section_name, section_df))

        logger.info(f"  - Section {i + 1}: {len(section_df)} rows (data only)")

    return sections