import pandas as pd
import os
import warnings
import re
import numpy as np

# --- Configuration ---

SOURCE_FOLDER = 'Downloaded files'
OUTPUT_FILE = 'summary_output_final.csv'

# Set to None to process ALL files in the folder
TEST_LIMIT = None 

# --- Helper Functions ---

def clean_money(value):
    """
    Cleans up messy numbers from Excel. 
    It handles things like '$ 1,000', '(500)' for negatives, and '-' for zero.
    """
    if pd.isna(value) or value == '': return 0.0
    if isinstance(value, (int, float)): return float(value)
    
    str_val = str(value).strip()
    # Accounting notation for zero
    if str_val in ['-', '–', '—', 'N/A', 'n/a']: return 0.0
    
    # Handle negative numbers in parentheses like (100)
    is_negative = False
    if '(' in str_val and ')' in str_val:
        is_negative = True
        str_val = str_val.replace('(', '').replace(')', '')
    
    # Remove everything that isn't a number or a decimal point
    clean_str = re.sub(r'[^\d.]', '', str_val)
    
    try:
        amount = float(clean_str)
        return -amount if is_negative else amount
    except ValueError:
        return 0.0

def find_rows_with_keywords(df, keywords):
    """
    Scans the dataframe to find which rows contain specific words.
    Useful because the row number changes from file to file.
    """
    matches = []
    # Convert the whole sheet to lowercase text for easy searching
    df_str = df.astype(str).apply(lambda x: x.str.lower())
    
    if isinstance(keywords, str): keywords = [keywords]
    keywords = [k.lower() for k in keywords]
    
    for idx, row in df_str.iterrows():
        # Combine all cells in the row into one string
        row_text = " ".join(row.values)
        for k in keywords:
            if k in row_text:
                matches.append(idx)
                break
    return matches

def extract_best_unit_count(df, row_indices):
    """
    Smart logic to grab the Unit Count.
    It looks at the row where we found "Total Units" and the few rows below it.
    Crucially, it ignores the number "9" (which is just the question number) 
    and years like "2025".
    """
    candidates = []
    for r_idx in row_indices:
        # Look at the matching row and the next 3 rows (just in case)
        for offset in range(4):
            if r_idx + offset >= len(df): continue
            row = df.iloc[r_idx + offset]
            
            for cell in row:
                val = clean_money(cell)
                
                # We expect projects to have between 5 and 6000 units
                if val >= 5 and val <= 6000 and float(val).is_integer():
                    
                    # Ignore the number 9 (it's the question label "9. Total Units")
                    if abs(val - 9) < 0.1: continue
                    
                    # Ignore years (we don't want 2024 or 2025 showing up as units)
                    if 2018 <= val <= 2030: continue
                    
                    candidates.append(val)
    
    # If we found multiple numbers, the largest one is likely the real unit count
    if candidates:
        return max(candidates)
    return 0.0

def extract_square_footage(df, row_indices):
    """
    Finds the square footage. 
    It ignores small numbers and huge outliers to ensure we get the building area.
    """
    candidates = []
    for r_idx in row_indices:
        for offset in range(3):
            if r_idx + offset >= len(df): continue
            row = df.iloc[r_idx + offset]
            
            for cell in row:
                val = clean_money(cell)
                # Valid SF range: 2,000 to 2,000,000
                if 2000 <= val <= 2000000:
                    candidates.append(val)
                    
    if candidates:
        return max(candidates)
    return 0.0

def calculate_hard_costs_robust(df):
    """
    Calculates Hard Costs. 
    First, it looks for a "Total" line. 
    If that's missing or zero, it manually sums up the individual cost items 
    (Site Work, Structures, etc.) to reconstruct the total.
    """
    # Option 1: Look for the Total line directly
    hc_keywords = ["Total New Construction Costs", "Total Rehabilitation Costs", "Total Hard Costs"]
    hc_rows = find_rows_with_keywords(df, hc_keywords)
    
    total_hc_line = 0.0
    if hc_rows:
        for r in hc_rows:
            # Grab the largest number in that row
            vals = [clean_money(x) for x in df.iloc[r] if clean_money(x) > 1000]
            if vals:
                total_hc_line += max(vals)
    
    # If we found a valid total, return it
    if total_hc_line > 0:
        return total_hc_line

    # Option 2: Fallback - Sum the parts manually
    components = [
        "Site Work", "Structures", "General Requirements", 
        "Contractor Overhead", "Contractor Profit", "Prevailing Wages"
    ]
    
    sum_hc = 0.0
    for comp in components:
        rows = find_rows_with_keywords(df, [comp])
        for r in rows:
            # Skip rows that say "Total" so we don't double count
            row_text = " ".join([str(x).lower() for x in df.iloc[r]])
            if "total" in row_text: continue 
            
            vals = [clean_money(x) for x in df.iloc[r] if clean_money(x) > 100]
            if vals:
                sum_hc += max(vals)
        
    return sum_hc

def extract_section_costs(df, start_marker_list, end_marker_list):
    """
    Extracts all line items between two section headers (e.g., Construction Start and End).
    """
    start_rows = find_rows_with_keywords(df, start_marker_list)
    end_rows = find_rows_with_keywords(df, end_marker_list)
    
    if not start_rows or not end_rows:
        return {}, 0.0
    
    s_idx = start_rows[0]
    # Find the first 'Total' line that appears AFTER the start header
    e_idx = next((e for e in end_rows if e > s_idx), None)
    
    if e_idx is None: return {}, 0.0
    
    # Get the official total from the sheet
    end_row_vals = [clean_money(x) for x in df.iloc[e_idx] if clean_money(x) > 100]
    official_total = max(end_row_vals) if end_row_vals else 0.0
    
    data = {}
    
    # Loop through the rows in between
    for i in range(s_idx + 1, e_idx):
        row = df.iloc[i]
        vals = [clean_money(x) for x in row]
        # Valid cost filter: > $100 and not a year
        valid_vals = [v for v in vals if abs(v) > 100 and v not in [2024, 2025]]
        
        if valid_vals:
            cost = max(valid_vals, key=abs)
            
            # Try to guess the label from the text in the row
            row_text_parts = [str(x) for x in row if pd.notna(x) and not str(x).strip().startswith(('$', '0', 'nan'))]
            if row_text_parts:
                label = max(row_text_parts, key=len) # Longest text is usually the label
                label = re.sub(r'[0-9$(),]', '', label).strip()
            else:
                label = "Other Cost"
            
            if len(label) < 3: label = "Other Cost"
            if "other" in label.lower(): label = "Other: " + label
            
            # Avoid duplicate keys
            if label in data: label = label + "_2"
            
            data[label] = cost
            
    return data, official_total

# --- Main Execution ---

if __name__ == "__main__":
    processed_rows = []
    # Silence Excel reading warnings
    warnings.filterwarnings('ignore')

    if not os.path.exists(SOURCE_FOLDER):
        print(f"Error: Folder '{SOURCE_FOLDER}' not found. Please check the path.")
        exit()

    # Get list of Excel files
    all_files = [f for f in os.listdir(SOURCE_FOLDER) if f.lower().endswith(('.xlsx', '.xlsm', '.xls'))]
    all_files.sort()
    
    # Apply limit if testing, otherwise process everything
    if TEST_LIMIT:
        files_to_process = all_files[:TEST_LIMIT]
    else:
        files_to_process = all_files

    print(f"Starting processing of {len(files_to_process)} files...")

    for i, filename in enumerate(files_to_process, 1):
        filepath = os.path.join(SOURCE_FOLDER, filename)
        print(f"[{i}/{len(files_to_process)}] Processing {filename}...")
        
        row_data = {'File Name': filename, 'Flag': []}
        
        try:
            xls = pd.ExcelFile(filepath)
            
            # ---------------------------
            # 1. Find Total Units
            # ---------------------------
            # Strategy: Look in 'Application' tab first, then others if missing
            sheets_to_check = []
            app_sheet = next((s for s in xls.sheet_names if "App" in s and "Check" not in s), None)
            if app_sheet: sheets_to_check.append(app_sheet)
            
            # Add all other sheets as backup
            for s in xls.sheet_names:
                if s not in sheets_to_check: sheets_to_check.append(s)
                
            units_found = 0
            for sheet in sheets_to_check:
                df = pd.read_excel(xls, sheet_name=sheet, header=None)
                # Synonyms for 'Total Units'
                keywords = ["total units", "total residential units", "total number of units", "unit count"]
                matches = find_rows_with_keywords(df, keywords)
                
                if matches:
                    units = extract_best_unit_count(df, matches)
                    if units > 0:
                        units_found = units
                        break # Found it, stop searching
            
            row_data['Total Units'] = units_found
            if units_found == 0: row_data['Flag'].append("Units Not Found")

            # ---------------------------
            # 2. Find Square Footage
            # ---------------------------
            sf_found = 0
            for sheet in sheets_to_check:
                df = pd.read_excel(xls, sheet_name=sheet, header=None)
                # Synonyms for SF
                keywords = ["total square footage", "residential sq", "gross building area", "net rentable area", "total s.f."]
                matches = find_rows_with_keywords(df, keywords)
                
                if matches:
                    sf = extract_square_footage(df, matches)
                    if sf > 0:
                        sf_found = sf
                        break
            
            row_data['Total SF'] = sf_found
            if sf_found == 0: row_data['Flag'].append("SF Not Found")

            # ---------------------------
            # 3. Financial Data
            # ---------------------------
            # Find the budget sheet
            budget_sheet = next((s for s in xls.sheet_names if "Source" in s and "Use" in s), None)
            if not budget_sheet: budget_sheet = next((s for s in xls.sheet_names if "Budget" in s), None)
            
            if budget_sheet:
                df_bud = pd.read_excel(xls, sheet_name=budget_sheet, header=None)
                
                # A. Construction Costs
                const_items, const_total = extract_section_costs(
                    df_bud, 
                    ["CONSTRUCTION INTEREST", "CONSTRUCTION FINANCING"], 
                    ["Total Construction Interest", "Total Construction Financing"]
                )
                row_data['Const Total (Sheet)'] = const_total
                row_data['Const Total (Calculated)'] = sum(const_items.values())
                
                # Map specific line items to columns
                row_data['Const Loan Interest'] = 0
                row_data['Const Origination Fee'] = 0
                details = []
                
                for k, v in const_items.items():
                    kl = k.lower()
                    if "interest" in kl and "construction" in kl: 
                        row_data['Const Loan Interest'] += v
                    elif "origination" in kl or "loan fee" in kl: 
                        row_data['Const Origination Fee'] += v
                    else: 
                        # Keep track of everything else in the details column
                        details.append(f"{k}:${int(v)}")
                row_data['Const Other Details'] = "; ".join(details)
                
                # B. Permanent Costs
                perm_items, perm_total = extract_section_costs(
                    df_bud, 
                    ["PERMANENT FINANCING", "PERMANENT SOURCES"], 
                    ["Total Permanent Financing", "Total Permanent Sources"]
                )
                row_data['Perm Total (Sheet)'] = perm_total
                row_data['Perm Total (Calculated)'] = sum(perm_items.values())
                
                # C. Hard Costs
                row_data['Hard Costs'] = calculate_hard_costs_robust(df_bud)
                if row_data['Hard Costs'] == 0: row_data['Flag'].append("Hard Costs 0")

            else:
                row_data['Flag'].append("Budget Tab Missing")

            # ---------------------------
            # 4. Final Calculations
            # ---------------------------
            combined = row_data.get('Const Total (Calculated)', 0) + row_data.get('Perm Total (Calculated)', 0)
            row_data['Combined Financing Costs'] = combined
            
            u = row_data.get('Total Units', 0)
            if u > 0: 
                row_data['Financing Cost per Unit'] = combined / u
            
            sf = row_data.get('Total SF', 0)
            if sf > 0: 
                row_data['Financing Cost per SF'] = combined / sf
            
            hc = row_data.get('Hard Costs', 0)
            if hc > 0: 
                row_data['% of Hard Costs'] = (combined / hc) * 100
            
            # Clean up flags
            row_data['Flag'] = "; ".join(row_data['Flag'])
            processed_rows.append(row_data)

        except Exception as e:
            print(f"  Error processing file: {e}")
            processed_rows.append({'File Name': filename, 'Flag': str(e)})

    # --- Export to CSV ---
    df_final = pd.DataFrame(processed_rows)

    # Set up the exact columns we want in the output
    standard_cols = [
        "File Name", "Flag",
        "Combined Financing Costs", "Financing Cost per Unit", "Financing Cost per SF", "% of Hard Costs",
        "Total Units", "Total SF", "Hard Costs",
        "Const Total (Sheet)", "Const Total (Calculated)", 
        "Const Loan Interest", "Const Origination Fee", "Const Other Details",
        "Perm Total (Sheet)", "Perm Total (Calculated)"
    ]
    
    # Fill in any missing columns with 0
    for c in standard_cols:
        if c not in df_final.columns: df_final[c] = 0

    df_final = df_final[standard_cols]
    df_final.to_csv(OUTPUT_FILE, index=False)

    print(f"\nSuccess! Summary saved to {OUTPUT_FILE}")