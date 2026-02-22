import pandas as pd
import json
import os
import glob
from datetime import time, datetime

def format_time_val(val):
    if pd.isna(val) or val == "" or str(val).strip().lower() == "nan":
        return "—"
    if isinstance(val, time):
        return val.strftime('%I:%M %p').lstrip('0')
    if isinstance(val, datetime):
        return val.strftime('%I:%M %p').lstrip('0')
    # Try to parse string time
    try:
        t_str = str(val).strip()
        if ":" in t_str:
            # Handle HH:MM:SS or HH:MM
            parts = t_str.split(':')
            h = int(parts[0])
            m = int(parts[1])
            t_obj = time(h, m)
            return t_obj.strftime('%I:%M %p').lstrip('0')
    except:
        pass
    return str(val).strip()

def process_file(excel_file):
    print(f"Reading {excel_file}...")
    try:
        xl = pd.ExcelFile(excel_file)
    except Exception as e:
        print(f"Error reading {excel_file}: {e}")
        return {}

    file_data = {}

    for sheet_name in xl.sheet_names:
        print(f"  Processing sheet: {sheet_name}")
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
        
        # Header Detection logic
        # Look in first 5 rows for keywords
        mapping = {}
        header_row_idx = -1
        for i in range(min(5, len(df))):
            row_vals = [str(x).lower().strip() for x in df.iloc[i].tolist()]
            if 'date' in row_vals and ('fajr' in " ".join(row_vals) or 'suhur' in " ".join(row_vals)):
                header_row_idx = i
                # Found header row
                for idx, val in enumerate(row_vals):
                    if 'date' in val: mapping['date'] = idx
                    if 'ramadan' in val: mapping['ramadan'] = idx
                    
                    # Suhur/Fajr Mapping
                    if 'fajr_18' in val or 'suhur_18' in val: mapping['fajr_18'] = idx
                    if 'fajr_na' in val or 'suhur_na' in val: mapping['fajr_na'] = idx
                    if 'fajr' in val and 'iqamah' in val: mapping['fajr_iqamah'] = idx
                    
                    if 'sunrise' in val: mapping['sunrise'] = idx
                    
                    if 'dhuhr' in val and ('start' in val or 'dhuhr' == val): mapping['zuhr'] = idx
                    if 'dhuhr' in val and 'iqamah' in val: mapping['zuhr_iqamah'] = idx
                    
                    # Asr Mapping
                    if 'asr' in val and 'standard' in val: mapping['asr_standard'] = idx
                    if 'asr' in val and 'hanafi' in val: mapping['asr_hanafi'] = idx
                    if 'hanafi' in val and 'asr' not in val: mapping['asr_hanafi'] = idx # Handle split labels
                    if 'asr' in val and 'iqamah' in val: mapping['asr_iqamah'] = idx
                    
                    if ('maghrib' in val or 'iftar' in val) and 'start' in val: mapping['maghrib'] = idx
                    elif ('maghrib' in val or 'iftar' in val) and 'iqamah' not in val and 'maghrib' not in mapping: mapping['maghrib'] = idx
                    if ('maghrib' in val or 'iftar' in val) and 'iqamah' in val: mapping['maghrib_iqamah'] = idx
                    
                    if 'isha' in val and ('start' in val or 'isha' == val): mapping['isha'] = idx
                    if 'isha' in val and 'iqamah' in val: mapping['isha_iqamah'] = idx
                break
        
        # Fallback to hardcoded legacy indices if no header detected
        if not mapping or ('fajr_18' not in mapping and 'fajr_na' not in mapping):
            print(f"    No modern header found, using legacy mapping for {sheet_name}")
            mapping = {
                'date': 0, 'fajr_18': 3, 'fajr_na': 3, 'fajr_iqamah': 4, 'sunrise': 5, 
                'zuhr': 6, 'zuhr_iqamah': 7, 'asr_standard': 9, 'asr_hanafi': 9, 'asr_iqamah': 10,
                'maghrib': 12, 'maghrib_iqamah': 13, 'isha': 14, 'isha_iqamah': 15
            }
        else:
            print(f"    Dynamic mapping detected: {list(mapping.keys())}")

        for i in range(header_row_idx + 1, len(df)):
            row = df.iloc[i]
            if len(row) <= mapping.get('date', 0): continue
            cell_value = row[mapping.get('date', 0)]
            if pd.isna(cell_value):
                continue
                
            try:
                if isinstance(cell_value, datetime):
                    cell_date = cell_value
                else:
                    cell_date = pd.to_datetime(cell_value)
                
                if pd.isna(cell_date):
                    continue
                    
                date_str = cell_date.strftime('%Y-%m-%d')
                
                ramadan_text = None
                if 'ramadan' in mapping:
                    r_val = str(row[mapping['ramadan']]).strip()
                    if r_val and r_val.lower() != 'nan' and any(c.isdigit() for c in r_val):
                        nums = ''.join(filter(str.isdigit, r_val))
                        if nums:
                            ramadan_text = f"Ramadan Day {int(nums)}"

                file_data[date_str] = {
                    "fajr_18": format_time_val(row[mapping.get('fajr_18', 0)]) if 'fajr_18' in mapping else "—",
                    "fajr_na": format_time_val(row[mapping.get('fajr_na', 0)]) if 'fajr_na' in mapping else "—",
                    "fajr_iqamah": format_time_val(row[mapping.get('fajr_iqamah', 4)]),
                    "sunrise": format_time_val(row[mapping.get('sunrise', 5)]),
                    "zuhr": format_time_val(row[mapping.get('zuhr', 6)]),
                    "zuhr_iqamah": format_time_val(row[mapping.get('zuhr_iqamah', 7)]),
                    "asr_standard": format_time_val(row[mapping.get('asr_standard', 0)]) if 'asr_standard' in mapping else "—",
                    "asr_hanafi": format_time_val(row[mapping.get('asr_hanafi', 0)]) if 'asr_hanafi' in mapping else "—",
                    "asr_iqamah": format_time_val(row[mapping.get('asr_iqamah', 10)]),
                    "maghrib": format_time_val(row[mapping.get('maghrib', 12)]),
                    "maghrib_iqamah": format_time_val(row[mapping.get('maghrib_iqamah', 13)]),
                    "isha": format_time_val(row[mapping.get('isha', 14)]),
                    "isha_iqamah": format_time_val(row[mapping.get('isha_iqamah', 15)]),
                    "ramadan": ramadan_text
                }
            except:
                continue
    return file_data

def run_conversion():
    output_dir = 'js'
    output_file = os.path.join(output_dir, 'prayerData.js')

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    all_combined_data = {}
    
    # Get all xlsx files
    files = glob.glob("*.xlsx")
    # Filter out temp files (starting with ~$)
    files = [f for f in files if not f.startswith('~$')]
    
    # Sort files by modification time (oldest to newest)
    # This ensures that the newest file's data overwrites older files if dates overlap
    files.sort(key=lambda x: os.path.getmtime(x))
    
    print(f"Found {len(files)} Excel files to process (ordered by modification time).")

    for f in files:
        file_results = process_file(f)
        if file_results:
            dates = sorted(file_results.keys())
            print(f"    File {f} covers: {dates[0]} to {dates[-1]}")
        
        # Group by Month for JSON structure compat
        for date_str, data in file_results.items():
            dt = datetime.strptime(date_str, '%Y-%m-%d')
            month_key = dt.strftime('%b') # 'Jan', 'Feb', etc.
            
            if month_key not in all_combined_data:
                all_combined_data[month_key] = {}
            
            # Overwrite or merge
            all_combined_data[month_key][date_str] = data

    # Sort results by date within each month
    for m in all_combined_data:
        all_combined_data[m] = dict(sorted(all_combined_data[m].items()))

    # Write as a JS variable
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("var prayerData = ")
        json.dump(all_combined_data, f, indent=2)
        f.write(";")
    
    print(f"\nSuccessfully generated {output_file} from {len(files)} files.")

if __name__ == "__main__":
    run_conversion()
