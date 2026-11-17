import pandas as pd
import os
from datetime import datetime
import sys
import traceback
import time
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

class TeeLogger:
    """Logger that writes to both console and file with immediate flushing."""
    def __init__(self, log_path):
        self.terminal = sys.stdout
        self.log = open(log_path, 'w', encoding='utf-8', buffering=1)
        self.start_time = time.time()
    
    def write(self, message):
        self.terminal.write(message)
        self.terminal.flush()
        self.log.write(message)
        self.log.flush()
    
    def flush(self):
        self.terminal.flush()
        self.log.flush()
    
    def close(self):
        self.log.close()

def print_log(message, level="INFO"):
    """Print with timestamp and flush immediately."""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    elapsed = time.time() - start_time if 'start_time' in globals() else 0
    formatted = f"[{timestamp}] [{level}] [{elapsed:.1f}s] {message}"
    print(formatted)
    sys.stdout.flush()

def print_separator(char="=", length=80):
    """Print a separator line."""
    print(char * length)
    sys.stdout.flush()

def print_progress_bar(current, total, prefix='', suffix='', length=50):
    """Print a progress bar with enhanced visual feedback."""
    if total == 0:
        return
    percent = 100 * (current / float(total))
    filled = int(length * current // total)
    bar = '█' * filled + '-' * (length - filled)
    print(f'\r{prefix} |{bar}| {percent:.1f}% {suffix}', end='', flush=True)
    if current == total:
        print()

def select_file(title, filetypes):
    """Open a file picker dialog."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=filetypes
    )
    root.destroy()
    return file_path

def wait_for_file(path, timeout=60):
    """Wait until file is accessible (not locked by another program)."""
    print_log(f"Checking file access: {os.path.basename(path)}")
    start = time.time()
    attempt = 0
    while True:
        try:
            with open(path, 'r+b') as f:
                print_log(f"✓ File access confirmed", "SUCCESS")
                return True
        except PermissionError:
            attempt += 1
            elapsed = time.time() - start
            if elapsed > timeout:
                raise PermissionError(f"File locked after {timeout}s: {os.path.basename(path)}")
            print_log(f"⚠ File locked (attempt {attempt}), retrying in 2s...", "WARNING")
            time.sleep(2)
        except FileNotFoundError:
            raise FileNotFoundError(f"File not found: {path}")

def auto_format_excel(file_path, packaging_levels, max_width=50):
    """Apply professional Excel formatting with color-coded packaging level blocks."""
    try:
        print_log("Applying Excel formatting...")
        wb = load_workbook(file_path)
        ws = wb.active

        # Color palette for each packaging level block
        level_colors = {
            'Consumer':  {"bg": "1F6B3B", "text": "FFFFFF"},  # Dark green
            'Customer':  {"bg": "1F4E79", "text": "FFFFFF"},  # Dark blue
            'Transport': {"bg": "7B2C2C", "text": "FFFFFF"},  # Dark red
        }
        default_header = {"bg": "366092", "text": "FFFFFF"}   # Default navy

        # Build a map: column index -> level name (or None for base columns)
        header_level_map = {}
        for col_idx, cell in enumerate(ws[1], start=1):
            col_name = str(cell.value) if cell.value else ''
            matched_level = None
            for level in packaging_levels:
                if col_name.startswith(f'{level}_'):
                    matched_level = level
                    break
            header_level_map[col_idx] = matched_level

        # Apply header formatting
        for col_idx, cell in enumerate(ws[1], start=1):
            level = header_level_map.get(col_idx)
            if level and level in level_colors:
                colors = level_colors[level]
            else:
                colors = default_header

            cell.font = Font(bold=True, color=colors["text"], size=11)
            cell.fill = PatternFill(
                start_color=colors["bg"],
                end_color=colors["bg"],
                fill_type="solid"
            )
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Freeze top row
        ws.freeze_panes = "A2"

        # Auto-adjust column widths
        for col in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = min(max_length + 3, max_width)
            ws.column_dimensions[col_letter].width = adjusted_width

        # Add filters to header row
        ws.auto_filter.ref = ws.dimensions

        wb.save(file_path)
        print_log("✓ Excel formatting completed", "SUCCESS")
    except Exception as e:
        print_log(f"⚠ Could not apply formatting: {e}", "WARNING")

def load_excel_optimized(file_path, sheet_name=0, usecols=None):
    """Load Excel with optimized settings for large files."""
    file_name = os.path.basename(file_path)
    file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
    
    print_separator()
    print_log(f"Loading: {file_name} ({file_size_mb:.2f} MB)")
    
    try:
        load_start = time.time()
        
        read_params = {
            'sheet_name': sheet_name,
            'engine': 'openpyxl',
            'dtype': str,
        }
        
        if usecols:
            read_params['usecols'] = usecols
        
        df = pd.read_excel(file_path, **read_params)
        
        load_time = time.time() - load_start
        print_log(f"✓ Loaded in {load_time:.1f}s", "SUCCESS")
        print_log(f"  Rows: {len(df):,} | Columns: {len(df.columns)}")
        print_log(f"  Memory: {df.memory_usage(deep=True).sum() / 1024**2:.2f} MB")
        
        return df
    except Exception as e:
        print_log(f"✗ Failed to load: {e}", "ERROR")
        raise

def normalize_dataframe(df):
    """Normalize DataFrame: strip whitespace and handle data types."""
    print_log("Normalizing DataFrame...")
    
    df.columns = df.columns.str.strip()
    
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace('nan', pd.NA)
    
    print_log(f"✓ Normalization complete", "SUCCESS")
    return df

def validate_columns(df, required_columns, file_description):
    """Validate presence of required columns with helpful error messages."""
    print_log(f"Validating columns in {file_description}...")
    missing_cols = [col for col in required_columns if col not in df.columns]
    
    if missing_cols:
        print_log(f"✗ Missing columns: {', '.join(missing_cols)}", "ERROR")
        print_log(f"  Available: {', '.join(df.columns.tolist()[:10])}", "INFO")
        raise ValueError(f"Missing required columns: {', '.join(missing_cols)}")
    
    print_log(f"✓ All {len(required_columns)} required columns present", "SUCCESS")

def search_master_by_material(material_code, master):
    """Search master file by Material with fallback to FG Code."""
    subset = master[master['Material'] == material_code]
    if not subset.empty:
        return subset
    
    subset = master[master['FG Code'] == material_code]
    if not subset.empty:
        return subset
    
    return pd.DataFrame()


# ──────────────────────────────────────────────────────────────────────────────
# UPDATED: process_codes_vectorized now separates output by packaging level
# Output columns follow the pattern:  Consumer_Plastic, Customer_Metal, etc.
# Three clear blocks per material: Consumer | Customer | Transport
# ──────────────────────────────────────────────────────────────────────────────
def process_codes_vectorized(material_codes, master, fixed_constituents, packaging_levels):
    """
    Process material codes and output one row per material, with separate
    column blocks for each EC Packaging Level (Consumer / Customer / Transport).

    Column naming convention:
        {Level}_{Constituent}              e.g.  Consumer_Plastic
        {Level}_Metals_without_Aluminum    e.g.  Customer_Metals_without_Aluminum
        {Level}_Aluminum                   e.g.  Transport_Aluminum
    """
    print_separator()
    print_log(f"Processing {len(material_codes):,} unique material codes...")
    print_log(f"Packaging levels: {', '.join(packaging_levels)}")
    process_start = time.time()

    result_rows  = []
    missing_codes = []

    total               = len(material_codes)
    checkpoint_interval = max(100, total // 20)   # update every ~5 %

    for idx, material_code in enumerate(material_codes, 1):

        # ── progress bar ──────────────────────────────────────────────────────
        if idx % checkpoint_interval == 0 or idx == total:
            elapsed = time.time() - process_start
            rate    = idx / elapsed if elapsed > 0 else 0
            eta     = (total - idx) / rate if rate > 0 else 0
            print_progress_bar(
                idx, total,
                prefix='Progress',
                suffix=f'{rate:.0f} codes/sec | ETA: {eta:.0f}s'
            )

        # ── look up material in master ─────────────────────────────────────────
        subset = search_master_by_material(material_code, master)

        if subset.empty:
            # ── material not found: empty row with zeros ───────────────────────
            missing_codes.append(material_code)
            row = {
                'Material':    material_code,
                'EAN':         '',
                'Code Title':  '',
                'Gross Weight': '',
                'Net Weight':   ''
            }
            for level in packaging_levels:
                for c in fixed_constituents:
                    row[f'{level}_{c}'] = 0.0
                row[f'{level}_Metals_without_Aluminum'] = 0.0
                row[f'{level}_Aluminum']                = 0.0
                row[f'{level}_% Recycled']               = ''

            row['% Recycled Content']   = ''
            row['Recycled Weight (kg)'] = ''

        else:
            # ── material found ─────────────────────────────────────────────────
            first_row = subset.iloc[0]

            # EAN: handle scientific notation
            ean_value = first_row.get('EAN/UPC', '')
            if ean_value and ean_value != '':
                try:
                    ean_value = str(int(float(ean_value)))
                except Exception:
                    ean_value = str(ean_value)

            # Base / product-level fields (same regardless of packaging level)
            row = {
                'Material':     material_code,
                'EAN':          ean_value,
                'Code Title':   first_row.get('Code Title',   ''),
                'Gross Weight': first_row.get('Gross Weight', ''),
                'Net Weight':   first_row.get('Net Weight',   '')
            }

            # ── per-level constituent weights ──────────────────────────────────
            for level in packaging_levels:
                level_mask   = subset['EC Packaging Level'].str.contains(
                    level, case=False, na=False
                )
                level_subset = subset[level_mask]

                if level_subset.empty:
                    # No rows for this level → zeros
                    for c in fixed_constituents:
                        row[f'{level}_{c}'] = 0.0
                    row[f'{level}_Metals_without_Aluminum'] = 0.0
                    row[f'{level}_Aluminum']                = 0.0
                    row[f'{level}_% Recycled']               = ''
                    continue

                # Sum Weight by Constituent for this level
                weight_sums = level_subset.groupby('Constituent')['Weight'].sum()

                for c in fixed_constituents:
                    row[f'{level}_{c}'] = float(weight_sums.get(c, 0.0))

                # Aluminum for this level
                alu_mask = level_subset['ConstituentCategory'] == 'Aluminum'
                row[f'{level}_Aluminum'] = float(
                    level_subset[alu_mask]['Weight'].sum()
                )

                # Metals without Aluminum for this level
                metal_no_alu_mask = (
                    (level_subset['Constituent'] == 'Metal') &
                    (level_subset['ConstituentCategory'] != 'Aluminum')
                )
                row[f'{level}_Metals_without_Aluminum'] = float(
                    level_subset[metal_no_alu_mask]['Weight'].sum()
                )

                # Recycled % for this level = recycled weight / total weight
                level_weight = level_subset['Weight'].sum()
                row[f'{level}_% Recycled'] = (
                    round(level_subset['Recycled WQF'].sum() / level_weight * 100, 2)
                    if level_weight > 0 else ''
                )

            # Recycled content: totals across all components and levels
            total_weight   = subset['Weight'].sum()
            total_recycled = subset['Recycled WQF'].sum()
            row['% Recycled Content']   = (
                round(total_recycled / total_weight * 100, 2) if total_weight > 0 else ''
            )
            row['Recycled Weight (kg)'] = total_recycled / 1000

        result_rows.append(row)

    process_time = time.time() - process_start
    print_log(f"✓ Processing complete in {process_time:.1f}s", "SUCCESS")
    print_log(
        f"  Matched: {len(material_codes) - len(missing_codes):,} | "
        f"Missing: {len(missing_codes):,}"
    )
    return result_rows, missing_codes


def main():
    global start_time
    start_time = time.time()
    
    print_separator("=", 80)
    print_log("EXCEL DATA PROCESSOR v10.0 - MULTI-LEVEL PACKAGING SPLIT", "INFO")
    print_separator("=", 80)
    
    log_path = None
    
    try:
        # Setup logging
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        folder    = os.path.dirname(os.path.abspath(__file__)) or os.getcwd()
        log_path  = os.path.join(folder, f"PROCESS_LOG_{timestamp}.txt")
        
        sys.stdout = TeeLogger(log_path)
        print_log(f"Working directory: {folder}")
        print_log(f"Log file: {os.path.basename(log_path)}")
        
        # ── File selection ─────────────────────────────────────────────────────
        print_separator()
        print_log("Please select INPUT file...")
        input_path = select_file(
            "Select INPUT file (with Material codes)",
            [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not input_path:
            raise ValueError("No input file selected")
        print_log(f"✓ Input: {os.path.basename(input_path)}", "SUCCESS")
        
        print_log("Please select MASTER file...")
        master_path = select_file(
            "Select MASTER file (green_dot_all_merged)",
            [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not master_path:
            raise ValueError("No master file selected")
        print_log(f"✓ Master: {os.path.basename(master_path)}", "SUCCESS")
        
        output_path = os.path.join(folder, f"RESULT_{timestamp}.xlsx")
        
        # ── File access check ──────────────────────────────────────────────────
        print_separator()
        wait_for_file(master_path)
        wait_for_file(input_path)
        
        # ── Load master ────────────────────────────────────────────────────────
        master = load_excel_optimized(master_path)
        master = normalize_dataframe(master)
        
        # ── Validate master columns ────────────────────────────────────────────
        print_separator()
        required_master_cols = [
            'FG Code', 'Material', 'Code Title', 'Gross Weight', 'Net Weight',
            'Constituent', 'Weight', 'WeightQuantityFactor', 'EC Packaging Level', 'ConstituentCategory',
            'EAN/UPC', '% Recycled Content', 'Recycled WQF', 'Recycled Weight (kg)'
        ]
        validate_columns(master, required_master_cols, "master file")
        
        # Convert Weight to numeric
        # Weight per sold unit
        print_log("Using WeightQuantityFactor (weight per sold unit)...")
        master['Weight'] = pd.to_numeric(master['WeightQuantityFactor'], errors='coerce').fillna(0)
        master['Recycled WQF'] = pd.to_numeric(master['Recycled WQF'], errors='coerce').fillna(0)

        # Remove leading zeros so 099350169839 and 99350169839 match
        print_log("Removing leading zeros from Material and FG Code...")
        master['Material'] = master['Material'].str.lstrip('0')
        master['FG Code']  = master['FG Code'].str.lstrip('0')

        # ── UPDATED: filter for ALL THREE packaging levels ─────────────────────
        packaging_levels = ['Consumer', 'Customer', 'Transport']

        print_separator()
        print_log(f"Filtering for: {', '.join(packaging_levels)} packaging levels...")
        original_count = len(master)

        mask = False
        for lvl in packaging_levels:
            mask = mask | master['EC Packaging Level'].str.contains(
                lvl, case=False, na=False
            )
        master = master[mask].copy()

        print_log(
            f"✓ Filtered: {len(master):,} rows "
            f"(removed {original_count - len(master):,})",
            "SUCCESS"
        )

        # Show breakdown per level for transparency
        for lvl in packaging_levels:
            count = master['EC Packaging Level'].str.contains(
                lvl, case=False, na=False
            ).sum()
            print_log(f"  {lvl:12s}: {count:,} rows")

        # ── Load input ─────────────────────────────────────────────────────────
        input_df = load_excel_optimized(input_path)
        validate_columns(input_df, ['Material'], "input file")
        input_df = normalize_dataframe(input_df)
        
        # ── Extract unique material codes ──────────────────────────────────────
        print_separator()
        print_log("Extracting unique Material codes...")
        material_codes = (
            input_df['Material']
            .dropna()
            .astype(str)
            .str.strip()
            .str.lstrip('0')
        )
        material_codes = material_codes[material_codes != ''].unique().tolist()
        print_log(f"✓ Found {len(material_codes):,} unique material codes", "SUCCESS")
        
        if len(material_codes) == 0:
            raise ValueError("No valid material codes found in input file!")
        
        # ── Fixed constituent types ────────────────────────────────────────────
        fixed_constituents = [
            'Metal', 'Plastic', 'Cartonboard', 'Glass',
            'Paper', 'Corrugated', 'Other', 'Natural'
        ]
        
        # ── Process codes with level split ─────────────────────────────────────
        result_rows, missing_codes = process_codes_vectorized(
            material_codes, master, fixed_constituents, packaging_levels
        )
        
        # ── Build output columns: base + level blocks + recycled ───────────────
        print_separator()
        print_log("Creating output DataFrame...")

        base_cols  = ['Material', 'EAN', 'Code Title', 'Gross Weight', 'Net Weight']
        level_cols = []
        for level in packaging_levels:
            for c in fixed_constituents:
                level_cols.append(f'{level}_{c}')
            level_cols.append(f'{level}_Metals_without_Aluminum')
            level_cols.append(f'{level}_Aluminum')
            level_cols.append(f'{level}_% Recycled')

        output_cols = base_cols + level_cols + ['% Recycled Content', 'Recycled Weight (kg)']

        result_df = pd.DataFrame(result_rows, columns=output_cols)
        print_log(
            f"✓ DataFrame: {len(result_df):,} rows × {len(result_df.columns)} cols",
            "SUCCESS"
        )
        
        # ── Save results ───────────────────────────────────────────────────────
        print_separator()
        print_log(f"Saving to: {os.path.basename(output_path)}")
        save_start = time.time()
        result_df.to_excel(output_path, index=False, engine='openpyxl')
        print_log(f"✓ Saved in {time.time() - save_start:.1f}s", "SUCCESS")
        
        # Apply color-coded formatting (one color per level block)
        auto_format_excel(output_path, packaging_levels)
        
        # ── Summary ────────────────────────────────────────────────────────────
        print_separator("=", 80)
        print_log("SUMMARY", "INFO")
        print_separator("=", 80)
        match_rate = (
            (len(material_codes) - len(missing_codes)) / len(material_codes) * 100
            if material_codes else 0
        )
        print_log(f"Packaging levels in output : {', '.join(packaging_levels)}")
        print_log(f"Total codes                : {len(material_codes):,}")
        print_log(f"Matched                    : {len(material_codes) - len(missing_codes):,} ({match_rate:.1f}%)")
        print_log(f"Missing                    : {len(missing_codes):,} ({100 - match_rate:.1f}%)")
        
        if missing_codes:
            print_separator()
            print_log("Missing codes (first 20):", "WARNING")
            for i, code in enumerate(missing_codes[:20], 1):
                print_log(f"  {i:2d}. {code}")
            if len(missing_codes) > 20:
                print_log(f"  ... and {len(missing_codes) - 20:,} more")
        
        total_time = time.time() - start_time
        print_separator("=", 80)
        print_log(f"✓ SUCCESS! Completed in {total_time:.1f}s", "SUCCESS")
        print_log(f"Output : {os.path.basename(output_path)}")
        print_log(f"Log    : {os.path.basename(log_path)}")
        print_separator("=", 80)
        
        sys.stdout.close()
        sys.stdout = sys.__stdout__
        
        print("\n✓ Process completed successfully!")
        print(f"Output: {output_path}")
        print(f"Log:    {log_path}")
        
    except Exception as e:
        print_separator("=", 80)
        print_log(f"✗ FATAL ERROR: {str(e)}", "ERROR")
        print_separator("=", 80)
        print_log(traceback.format_exc(), "ERROR")
        
        if hasattr(sys.stdout, 'close'):
            sys.stdout.close()
            sys.stdout = sys.__stdout__
        
        print("\n✗ Process failed! Check log for details.")
        if log_path:
            print(f"Log: {log_path}")
        
        input("\nPress Enter to exit...")
        sys.exit(1)
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
