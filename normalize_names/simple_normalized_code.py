import re

import numpy as np
import pandas as pd


def clean_name(name):
    """Clean business terms and titles from name"""
    patterns = {
        "titles": r"\b(Dr|Mr|Mrs|Ms|Prof|CEO|CFO|Sr\.Eng)\b\.?",
        "business": r"\b(Inc|Ltd|Co|Corp|LLC|DBA|C/O)\b",
        "suffix": r",?\s+(Jr|Sr|III|IV|II)\.?$",
    }
    name = re.sub("|".join(patterns.values()), "", name, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", name).strip(" ,&")


def extract_components(name):
    """Extract first and last names from cleaned name"""
    # Handle multi-person names
    parts = re.split(r"\s+(?:and|&)\s+", name, flags=re.IGNORECASE)
    parts = [p.strip() for p in parts if p.strip()]

    if len(parts) > 1:
        # Check for shared last name
        last_names = [p.split()[-1] for p in parts]
        if len(set(last_names)) == 1:
            return parts[0].split()[0], last_names[0]

        # Fallback to first valid name pair
        valid = next((p for p in parts if len(p.split()) >= 2), None)
        if valid:
            parts = valid.split()

    # Default single name handling
    parts = name.split()
    return (parts[0], parts[-1]) if len(parts) >= 2 else (name, "")


def normalize_name(full_name):
    """Main normalization function"""
    if not isinstance(full_name, str) or not full_name.strip():
        return "", "", ""

    cleaned = clean_name(full_name)
    first, last = extract_components(cleaned)

    # Extract suffix separately
    suffix_match = re.search(r"\b(Jr|Sr|III|IV|II)\b", cleaned, re.IGNORECASE)
    suffix = suffix_match.group(0) if suffix_match else ""

    # Clean special characters while preserving hyphens
    first = re.sub(r"[^a-zA-Z-]", "", first)
    last = re.sub(r"[^a-zA-Z-]", "", last)

    return first.strip(), last.strip(), suffix


def process_excel_file(input_path, sheet_name):
    """Process Excel file with name normalization"""
    try:
        df = pd.read_excel(input_path, sheet_name=sheet_name)
        df[["FirstName", "LastName", "Suffix"]] = df["Client_Name_AMS"].apply(
            lambda x: pd.Series(normalize_name(str(x)))
        )
        df["Full_Lastname"] = (
            df["LastName"] + " " + df["Suffix"].replace("", np.nan).fillna("")
        )
        return df
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return None
