# Libraries
import requests
import pandas as pd
import re
import os
from datetime import datetime, timedelta
from difflib import get_close_matches


# =====================================================================
# Configuration (input, output, column names, etc.)
INPUT_FILE = "Template.xlsx"
OUTPUT_FILE = "Milestone1.csv"

COLUMN_NAME_WITH_SITES = "Wastewater Site"   
COLUMN_MAPCHART_ID = "Mapchart Id"
COLUMN_STATE_USA = "STATE"
COLUMN_COUNTY_NAME = "USA County Name"

TARGET_METRIC = "gc_g_dry_weight_pmmov"
FUZZY_CUTOFF = 0.6

# Special cases that were not perfectly matched, so need to make to edge cases.
SITE_OVERRIDES = {
    "Palo Alto (optional San Jose/Sunnyvale/Gilroy)": "Palo Alto",
    "Utoy Creek, Atlanta (Johns Creek/College Park/RM Clayton/South River/Big Creek/Little River)": "Utoy Creek, Atlanta",
}
SITE_NEEDS_MANUAL_REVIEW = {
    "Eastern/Northwest/South/Southwest": (
        "No primary name and no city/county/state anchor. Likely a 4-plant system."
    ),
}
# =====================================================================

def find_column(df, target_name):
    target_norm = re.sub(r"\s+", " ", str(target_name)).strip().lower()
    for col in df.columns:
        if re.sub(r"\s+", " ", str(col)).strip().lower() == target_norm:
            return col
    raise KeyError(f"Could not find a column matching '{target_name}'. Available: {list(df.columns)}")

def fetch_master_index():
    print("Initializing engine and fetching API master index...")
    # This was the Request URL I found for the API, not GraphQL
    url = "https://storage.googleapis.com/wastewater-dev-data/json/plants.json"
    response = requests.get(url)
    response.raise_for_status()
    raw_data = response.json()

    if isinstance(raw_data, dict):
        plants = next((val for val in raw_data.values() if isinstance(val, list)), [])
        if not plants and all(isinstance(v, dict) for v in raw_data.values()):
            plants = list(raw_data.values())
        return plants
    return raw_data

def normalize_name(name):
    n = str(name).lower().strip()
    n = re.sub(r"[,.]", " ", n)
    n = re.sub(r"\b(city of|town of|village of|wastewater treatment (facility|plant)|water reclamation plant|wwtp|wrp)\b", " ", n)
    return re.sub(r"\s+", " ", n).strip()

def build_fuzzy_lookup(plants_list):
    lookup = {}
    for plant in plants_list:
        if isinstance(plant, dict):
            uid = plant.get("uid") or str(plant.get("id", ""))[:8]
            for raw in (plant.get("name", ""), plant.get("site_name", "")):
                raw = str(raw).strip()
                if not raw: continue
                lookup[raw.lower()] = uid
                lookup[normalize_name(raw)] = uid
    return lookup

def parse_site_entry(raw_site):
    m = re.match(r"^(.*?)\s*\(\s*optional:\s*(.*?)\s*\)\s*$", str(raw_site), flags=re.IGNORECASE)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return str(raw_site).strip(), None

def match_site(name, plant_lookup):
    clean = name.lower().strip()
    norm = normalize_name(name)

    if norm in plant_lookup: return plant_lookup[norm], norm
    if clean in plant_lookup: return plant_lookup[clean], clean

    for key, uid in plant_lookup.items():
        if clean in key or key in clean or norm in key or key in norm:
            return uid, key

    candidates = get_close_matches(norm, plant_lookup.keys(), n=1, cutoff=FUZZY_CUTOFF)
    if candidates: return plant_lookup[candidates[0]], candidates[0]

    return None, None

def main():
    # 1. Load the Client's Target List
    try:
        # Added absolute order index to guarantee rows never shift
        df_input = pd.read_excel(INPUT_FILE)
        df_input['_input_order'] = range(len(df_input))
        
        site_col = find_column(df_input, COLUMN_NAME_WITH_SITES)

        site_rows = []
        for position, (_, row) in enumerate(df_input.iterrows()):
            site_val = row[site_col]
            if pd.isna(site_val):
                continue
            site_rows.append({
                "raw_site": site_val,
                "_input_order": position,
            })
        print(f"Loaded {len(df_input)} total rows ({len(site_rows)} valid sites) from {INPUT_FILE}.")
    except FileNotFoundError:
        print(f"CRITICAL ERROR: {INPUT_FILE} not found. Exiting.")
        return

    # 2. API Engine
    plants_list = fetch_master_index()
    plant_lookup = build_fuzzy_lookup(plants_list)

    master_data = []
    unmatched = []
    needs_review = []

    # 3. Extraction Loop
    print("\nStarting data extraction...")
    for entry in site_rows:
        raw_site = entry["raw_site"]

        if raw_site in SITE_NEEDS_MANUAL_REVIEW:
            print(f"  [?] NEEDS REVIEW: '{raw_site}'")
            needs_review.append(raw_site)
            continue

        if raw_site in SITE_OVERRIDES:
            primary, optional = SITE_OVERRIDES[raw_site], None
        else:
            primary, optional = parse_site_entry(raw_site)

        matched_uid, matched_key = match_site(primary, plant_lookup)
        used_fallback = False

        if not matched_uid and optional:
            matched_uid, matched_key = match_site(optional, plant_lookup)
            used_fallback = True

        if not matched_uid:
            print(f"  [X] FAILED: Could not match '{raw_site}'")
            unmatched.append(raw_site)
            continue

        used_name = optional if used_fallback else primary
        print(f"  [\u2713] SUCCESS: Matched '{raw_site}' -> API UID: {matched_uid}")

        # Fetch payload
        data_url = f"https://storage.googleapis.com/wastewater-dev-data/json/{matched_uid}.json"
        try:
            plant_data = requests.get(data_url).json()
        except Exception:
            print(f"  [X] FAILED: Could not download JSON for UID {matched_uid}.")
            continue

        # Calculate the dynamic 9-3 calendar day window
        run_date = datetime.now()
        target_start = (run_date - timedelta(days=9)).strftime('%Y-%m-%d')
        target_end = (run_date - timedelta(days=3)).strftime('%Y-%m-%d')

        target_samples = []
        for sample in plant_data.get("samples", []):
            c_date = sample.get("collection_date", "")
            # Only keep samples within the exact calendar window
            if target_start <= c_date <= target_end:
                target_samples.append(sample)

        for sample in target_samples:
            row = {
                "_input_order": entry["_input_order"],
                "collection_date": sample.get("collection_date")  # Capture the date
            }
            targets = sample.get("targets", {})
            for pathogen_name, metrics in targets.items():
                # Multiply by 1,000,000
                if not isinstance(metrics, dict):
                    metric_value = 0
                else:
                    raw_value = metrics.get(TARGET_METRIC, 0)
                    if raw_value is None or (isinstance(raw_value, float) and pd.isna(raw_value)):
                        metric_value = 0
                    else:
                        try:
                            metric_value = float(raw_value)
                        except (TypeError, ValueError):
                            metric_value = 0

                row[pathogen_name] = metric_value * 1000000

            master_data.append(row)

    # 4. Format, Aggregate, and Pivot the Data
    print("\nMerging and pivoting data back to original template...")
    if master_data:
        df_master = pd.DataFrame(master_data)
        
        # Ensure collection_date is a clean string for column naming
        df_master['collection_date'] = pd.to_datetime(df_master['collection_date']).dt.strftime('%Y-%m-%d')
        
        # Pathogens are everything except the index and date columns
        pathogen_cols = [c for c in df_master.columns if c not in ["_input_order", "collection_date"]]
        df_master[pathogen_cols] = df_master[pathogen_cols].fillna(0)

        # Rename N Gene to COVID-19 early so it appears cleanly in the new column headers
        if 'N Gene' in df_master.columns:
            df_master = df_master.rename(columns={'N Gene': 'COVID-19'})
            pathogen_cols = [c if c != 'N Gene' else 'COVID-19' for c in pathogen_cols]
        
        # Step 4a: Spatial Average (Group by County AND Date)
        df_daily_avg = df_master.groupby(['_input_order', 'collection_date'])[pathogen_cols].mean().reset_index()
        
        # Step 4b: Pivot into Time-Series Wide Format
        df_wide = df_daily_avg.pivot_table(index='_input_order', columns='collection_date', values=pathogen_cols)
        
        # Step 4c: Sort Columns by Pathogen, then Date
        # The pivot_table defaults to Pathogen as the top level, Date as the second.
        df_wide = df_wide.sort_index(axis=1)
        
        # Flatten the MultiIndex columns (e.g., 'COVID-19_2026-06-03')
        flattened_cols = [f"{pathogen}_{date}" for pathogen, date in df_wide.columns]
        df_wide.columns = flattened_cols
        df_wide = df_wide.reset_index()
            
        # Left Join time-series data onto the pristine MapChart template
        df_final = pd.merge(df_input, df_wide, on='_input_order', how='left')
        generated_columns = len(flattened_cols)
    else:
        df_final = df_input
        generated_columns = 0

    # Clean up temp column
    df_final = df_final.drop(columns=["_input_order"])
    # Format to prevent scientific notation or heavy rounding
    pd.set_option("display.float_format", lambda x: f"{x:.10f}")

    # 5. Export CSV
    df_final.to_csv(OUTPUT_FILE, index=False, encoding='utf-8-sig')
    
    print(f"\nExtraction complete!")
    print(f"Total rows preserved: {len(df_final)}")
    if master_data:
        print(f"Total new time-series columns generated: {generated_columns}")
    print(f"Saved deliverable to: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()