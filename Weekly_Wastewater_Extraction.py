# Libraries
import requests
import pandas as pd
import re
import os
import smtplib
from email.message import EmailMessage
from datetime import datetime, timedelta
from difflib import get_close_matches

# =====================================================================
# Configuration (input, output, column names, etc.)
INPUT_FILE = "Template.xlsx"
OUTPUT_FILE = "Weekly_Wastewater_Report.csv"

COLUMN_NAME_WITH_SITES = "Wastewater Site"   
COLUMN_MAPCHART_ID = "Mapchart Id"
COLUMN_STATE_USA = "STATE"
COLUMN_COUNTY_NAME = "USA County Name"

TARGET_METRIC = "gc_g_dry_weight_pmmov"
FUZZY_CUTOFF = 0.6

# Special cases mapped directly to arrays of target facilities to solve Claire's email requests safely
SITE_OVERRIDES = {
    "Palo Alto (optional San Jose/Sunnyvale/Gilroy)": ["Palo Alto"],
    "Utoy Creek, Atlanta (Johns Creek/College Park/RM Clayton/South River/Big Creek/Little River)": ["Utoy Creek, Atlanta"],
    "Newark": ["Passaic Valley"],
    "Santa Cruz": ["Santa Cruz City"],
    "Eastern/Northwest/South/Southwest": ["Eastern Orange", "Northwest Orange", "South Orange", "Southwest Orange"],
    "Wolcott, Kaw Point, P20 Kansas City": ["Wolcott", "Kaw Point", "P20 Kansas City"],
    "San Diego": ["Point Loma"],
    "Honolulu": ["Sand Island", "Honouliuli"] 
}

SITE_NEEDS_MANUAL_REVIEW = {}
# =====================================================================

def find_column(df, target_name):
    target_norm = re.sub(r"\s+", " ", str(target_name)).strip().lower()
    for col in df.columns:
        if re.sub(r"\s+", " ", str(col)).strip().lower() == target_norm:
            return col
    raise KeyError(f"Could not find a column matching '{target_name}'.")

def fetch_master_index():
    print("Initializing engine and fetching API master index...")
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

def get_target_sites(raw_site):
    """Safely extracts a list of target sites to query without breaking real facility names."""
    raw = str(raw_site).strip()
    
    # 1. Check if it is a hardcoded override (solves Claire's specific edge cases)
    if raw in SITE_OVERRIDES:
        mapped = SITE_OVERRIDES[raw]
        return mapped if isinstance(mapped, list) else [mapped]
    
    # 2. Check for the "(optional: X)" format (like Marin County)
    m = re.match(r"^(.*?)\s*\(\s*optional:\s*(.*?)\s*\)\s*$", raw, flags=re.IGNORECASE)
    if m:
        return [m.group(1).strip(), m.group(2).strip()]
        
    # 3. Default to exactly what is in the cell
    return [raw]

def fetch_cdc_data(start_date, end_date, df_input):
    """PHASE 2: Extracts CDC Socrata Data for WVAL and takes the maximum value per county."""
    print("\nStarting Phase 2: Fetching CDC WVAL API data...")
    
    url = "https://data.cdc.gov/resource/atcp-73re.json?$limit=50000&$order=week_end DESC"
    
    try:
        response = requests.get(url)
        response.raise_for_status()
        df_cdc = pd.DataFrame(response.json())
        
        if df_cdc.empty:
            return []
            
        df_cdc['week_end'] = pd.to_datetime(df_cdc['week_end'], errors='coerce').dt.strftime('%Y-%m-%d')
        mask_date = (df_cdc['week_end'] >= start_date) & (df_cdc['week_end'] <= end_date)
        df_cdc = df_cdc[mask_date].copy()
        
        if df_cdc.empty:
            return []

        target_map = {
            'sars-cov-2': 'COVID-19',
            'influenza a': 'Influenza A',
            'rsv': 'RSV'
        }
        
        df_cdc['mapped_pathogen'] = df_cdc['pathogen_target'].str.lower().map(target_map)
        df_cdc = df_cdc.dropna(subset=['mapped_pathogen'])
        
        df_cdc['site_wval'] = pd.to_numeric(df_cdc['site_wval'], errors='coerce')
        df_cdc = df_cdc.dropna(subset=['site_wval'])
        
        # Drop historical records within the window so we only calculate using the newest WVAL per site
        df_cdc = df_cdc.drop_duplicates(subset=['site', 'mapped_pathogen'], keep='first')
        
        cdc_records = []
        county_col = find_column(df_input, COLUMN_COUNTY_NAME)
        state_col = find_column(df_input, COLUMN_STATE_USA)
        mapchart_col = find_column(df_input, COLUMN_MAPCHART_ID)
        
        input_lookups = []
        for _, input_row in df_input.iterrows():
            input_lookups.append({
                '_input_order': input_row['_input_order'],
                'county': str(input_row[county_col]).lower(),
                'state': str(input_row[state_col]).lower(),
                'mapchart': str(input_row[mapchart_col]).lower()
            })
            
        matched_data = []
        cdc_fast_rows = df_cdc[['counties_served', 'state_territory', 'mapped_pathogen', 'site_wval', 'site']].to_dict('records')
        
        state_abbr = {"alaska": "ak", "california": "ca", "texas": "tx", "alabama": "al", "arkansas": "ar", "arizona": "az", "colorado": "co", "connecticut": "ct", "delaware": "de", "florida": "fl", "georgia": "ga", "hawaii": "hi", "iowa": "ia", "idaho": "id", "illinois": "il", "indiana": "in", "kansas": "ks", "kentucky": "ky", "louisiana": "la", "massachusetts": "ma", "maryland": "md", "maine": "me", "michigan": "mi", "minnesota": "mn"} 
        
        for cdc_row in cdc_fast_rows:
            c_county = str(cdc_row.get('counties_served', '')).lower()
            c_county = c_county.replace(" county", "").replace(" parish", "").replace(" city", "").replace(" borough", "").replace(" municipality", "").strip()
            c_state = str(cdc_row.get('state_territory', '')).lower()
            c_state_abbr = state_abbr.get(c_state, "xxx")
            
            patho = cdc_row['mapped_pathogen']
            wval = cdc_row['site_wval']
            
            for item in input_lookups:
                input_id = item['_input_order']
                
                # Combine her columns into one string to search across all of them
                input_combined = f"{item['state']} {item['county']} {item['mapchart']}"
                
                # If the CDC county is in her combined text, AND the CDC state (or abbreviation) is in it, it's a match
                if (c_county in input_combined) and (c_state in input_combined or c_state_abbr in input_combined):
                    matched_data.append({
                        '_input_order': input_id,
                        'pathogen_col': f"CDC_Latest_{patho}_WVAL",
                        'wval': wval
                    })
                    
        if not matched_data:
            return []
            
        df_matches = pd.DataFrame(matched_data)
        
        # Group by the input row and pathogen, and take the maximum WVAL
        max_wvals = df_matches.groupby(['_input_order', 'pathogen_col'])['wval'].max().reset_index()
        
        for _, row in max_wvals.iterrows():
            cdc_records.append({
                '_input_order': row['_input_order'],
                'pathogen_col': row['pathogen_col'],
                'val': float(row['wval'])
            })
                        
        print(f"  [\u2713] CDC Integration Success: Pulled WVALs and isolated max values.")
        return cdc_records
        
    except Exception as e:
        print(f"  [X] FAILED: CDC Extraction error: {e}")
        return []

def send_email(csv_file_path):
    print("\nDrafting automated email to client...")
    bot_email = os.environ.get('SENDER_EMAIL')
    bot_password = os.environ.get('SENDER_PASSWORD')
    
    if not bot_email or not bot_password:
        print("  [!] Email credentials missing. Ensure SENDER_EMAIL and SENDER_PASSWORD are set in GitHub Secrets.")
        return

    msg = EmailMessage()
    msg['Subject'] = 'Automated Weekly Wastewater Data'
    msg['From'] = bot_email
    msg['To'] = 'claire.safrai@gmail.com' 
   
    
    msg.set_content("Hi Claire,\n\nAttached is the weekly automated wastewater dataset.\n\nBest,\nAutomated Bot")

    with open(csv_file_path, 'rb') as f:
        msg.add_attachment(f.read(), maintype='text', subtype='csv', filename=os.path.basename(csv_file_path))

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(bot_email, bot_password)
            smtp.send_message(msg)
        print("  [\u2713] Email sent successfully!")
    except Exception as e:
        print(f"  [X] FAILED: Could not send email. Error: {e}")

def main():
    try:
        df_input = pd.read_excel(INPUT_FILE)
        
        # FIX: Forward fill the state column to overwrite the visual blanks with actual state names
        state_col = find_column(df_input, COLUMN_STATE_USA)
        df_input[state_col] = df_input[state_col].ffill()
        
        df_input['_input_order'] = range(len(df_input))
        site_col = find_column(df_input, COLUMN_NAME_WITH_SITES)

        site_rows = []
        for position, (_, row) in enumerate(df_input.iterrows()):
            site_val = row[site_col]
            if pd.isna(site_val): continue
            site_rows.append({"raw_site": site_val, "_input_order": position})
    except FileNotFoundError:
        return

    # REVERTED: Back to the original 9-day WWS timeframe to preserve formatting
    run_date = datetime.now()
    target_start = (run_date - timedelta(days=28)).strftime('%Y-%m-%d')
    target_end = run_date.strftime('%Y-%m-%d')

    plants_list = fetch_master_index()
    plant_lookup = build_fuzzy_lookup(plants_list)

    master_data = []

    wws_name_map = {
        'N Gene': 'COVID-19',
        'InfA': 'Influenza A',
        'InfB': 'Influenza B'
    }

    print("\nStarting WastewaterScan data extraction (with Max Value logic)...")
    for entry in site_rows:
        raw_site = entry["raw_site"]

        if raw_site in SITE_NEEDS_MANUAL_REVIEW: continue

        # Parse all potential sites from the cell safely
        site_names = get_target_sites(raw_site)

        # Match all possible UIDs for this row
        matched_uids = []
        for name in site_names:
            uid, _ = match_site(name, plant_lookup)
            if uid and uid not in matched_uids:
                matched_uids.append(uid)

        if not matched_uids: continue

        # Fetch data for ALL matched UIDs in this county
        for matched_uid in matched_uids:
            data_url = f"https://storage.googleapis.com/wastewater-dev-data/json/{matched_uid}.json"
            try:
                plant_data = requests.get(data_url).json()
            except Exception:
                continue

            target_samples = [s for s in plant_data.get("samples", []) if target_start <= s.get("collection_date", "") <= target_end]

            for sample in target_samples:
                row = {"_input_order": entry["_input_order"], "collection_date": sample.get("collection_date")}
                targets = sample.get("targets", {})
                for pathogen_name, metrics in targets.items():
                    if isinstance(metrics, dict):
                        raw_value = metrics.get(TARGET_METRIC)
                        if raw_value is not None and pd.notna(raw_value):
                            try:
                                clean_patho = wws_name_map.get(pathogen_name, pathogen_name)
                                row[f"WWS_{clean_patho}"] = float(raw_value) * 1000000
                            except (TypeError, ValueError):
                                pass
                master_data.append(row)

    if master_data:
        df_wws = pd.DataFrame(master_data)
        df_wws['collection_date'] = pd.to_datetime(df_wws['collection_date']).dt.strftime('%Y-%m-%d')
        
        pathogen_cols = [c for c in df_wws.columns if c not in ["_input_order", "collection_date"]]
        
        # Take the .max() across all matched plants in the county for a given date
        df_daily_agg = df_wws.groupby(['_input_order', 'collection_date'])[pathogen_cols].max().reset_index()
        
        df_wide = df_daily_agg.pivot_table(index='_input_order', columns='collection_date', values=pathogen_cols)
        df_wide = df_wide.sort_index(axis=1)
        
        flattened_cols = [f"{pathogen}_{date}" for pathogen, date in df_wide.columns]
        df_wide.columns = flattened_cols
        df_wide = df_wide.reset_index()
            
        df_final = pd.merge(df_input, df_wide, on='_input_order', how='left')
    else:
        df_final = df_input

    # Independent CDC Integration (Kept at 90 days to prevent the Juneau 0s bug)
    cdc_start = (run_date - timedelta(days=90)).strftime('%Y-%m-%d')
    cdc_records = fetch_cdc_data(cdc_start, target_end, df_input)
    
    expected_cdc_cols = ["CDC_Latest_COVID-19_WVAL", "CDC_Latest_Influenza A_WVAL", "CDC_Latest_RSV_WVAL"]

    if cdc_records:
        df_cdc_flat = pd.DataFrame(cdc_records).pivot(index='_input_order', columns='pathogen_col', values='val').reset_index()
        df_final = pd.merge(df_final, df_cdc_flat, on='_input_order', how='left')
    else:
        print("  [!] No CDC data found for this period. Generating blank columns.")

    # Force the columns to exist even if the API returned absolutely nothing
    for col in expected_cdc_cols:
        if col not in df_final.columns:
            df_final[col] = 0
            
    # Explicitly fill NaNs for CDC columns with 0
    actual_cdc_cols = [c for c in df_final.columns if c.startswith("CDC_Latest_")]
    df_final[actual_cdc_cols] = df_final[actual_cdc_cols].fillna(0)

    # Drop administrative columns
    df_final = df_final.drop(columns=["_input_order", "COLOR", "Color"], errors="ignore")
    
    # NEW SORTING LOGIC: Pin CDC columns to the front, followed by WWS
    cols = list(df_final.columns)
    cdc_cols = sorted([c for c in cols if c.startswith("CDC_Latest_")])
    wws_cols = sorted([c for c in cols if c.startswith("WWS_")])
    base_cols = [c for c in cols if c not in cdc_cols and c not in wws_cols]
    
    df_final = df_final[base_cols + cdc_cols + wws_cols]

    pd.set_option("display.float_format", lambda x: f"{x:.10f}")
    df_final.to_csv(OUTPUT_FILE, index=False, encoding='utf-8-sig')
    print(f"\nExtraction complete! Saved to {OUTPUT_FILE}")

    send_email(OUTPUT_FILE)

if __name__ == "__main__":
    main()