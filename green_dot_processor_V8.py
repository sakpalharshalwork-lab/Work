import pandas as pd
import os
import logging
from datetime import datetime
import time
import tkinter as tk
from tkinter import filedialog, messagebox

class GreenDotProcessor:
    """Enhanced Green Dot Report Processor - Optimized for Speed"""
    
    def __init__(self):
        self.timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self.folder = None
        self.env_report_path = None
        self.fg_to_spec_path = None
        self.output_path = None
        self.log_file = None
        self.cp_dict_percent = {}
        self.fg_dict = {}
        
    def setup_logging(self):
        """Initialize logging configuration"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(self.log_file, mode='w', encoding='utf-8'),
                logging.StreamHandler()
            ],
            force=True
        )
        logging.info("="*60)
        logging.info("Green Dot Processor Started")
        logging.info(f"Timestamp: {self.timestamp}")
        logging.info("="*60)
    
    def select_files(self) -> bool:
        """File picker GUI for selecting input files and folder"""
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        
        try:
            # Select folder
            messagebox.showinfo("Step 1", "Select the folder containing Green Dot Report files")
            self.folder = filedialog.askdirectory(title="Select Green Dot Files Folder")
            if not self.folder:
                messagebox.showwarning("Cancelled", "No folder selected. Exiting.")
                return False
            
            # Select Environmental Report
            messagebox.showinfo("Step 2", "Select the Environmental_Report.xlsx file")
            self.env_report_path = filedialog.askopenfilename(
                title="Select Environmental Report",
                initialdir=self.folder,
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            if not self.env_report_path:
                messagebox.showwarning("Cancelled", "No Environmental Report selected. Exiting.")
                return False
            
            # Select FG to Spec file
            messagebox.showinfo("Step 3", "Select the FGtoSpec v2.xlsx file")
            self.fg_to_spec_path = filedialog.askopenfilename(
                title="Select FG to Spec Mapping",
                initialdir=self.folder,
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            if not self.fg_to_spec_path:
                messagebox.showwarning("Cancelled", "No FG to Spec file selected. Exiting.")
                return False
            
            # Set output paths
            self.output_path = os.path.join(self.folder, f"green_dot_all_merged_{self.timestamp}.xlsx")
            self.log_file = os.path.join(self.folder, f"green_dot_all_log_{self.timestamp}.log")
            
            messagebox.showinfo("Success", f"Files selected!\n\nOutput: {self.output_path}")
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"Error during file selection: {str(e)}")
            return False
        finally:
            root.destroy()
    
    def load_environmental_report(self):
        """Load Environmental Report - optimized"""
        try:
            logging.info(f"Loading Environmental Report...")
            df = pd.read_excel(self.env_report_path, engine="openpyxl")
            
            # Build normalized keys WITHOUT writing back into df (avoids dtype clash)
            cp_codes = (
                df.iloc[:, 0]
                .astype(str)
                .str.strip()
                .str.upper()
                .str.replace(r"\.0$", "", regex=True)
                .str.lstrip("0")
            )
            recycled_vals = pd.to_numeric(df.iloc[:, 6], errors="coerce").fillna(0).clip(0, 100)
            
            self.cp_dict_percent = dict(zip(cp_codes, recycled_vals))
            
            logging.info(f"Env rows read: {len(cp_codes):,} | Unique CP codes: {cp_codes.nunique():,}")
            logging.info(f"✅ Loaded {len(self.cp_dict_percent)} CP codes")
            
        except Exception as e:
            logging.exception("❌ Failed to load Environmental Report")
            self.cp_dict_percent = {}
    
    def load_fg_to_spec(self):
        """Load FG to Spec mapping - optimized"""
        try:
            logging.info(f"Loading FG to Spec mapping...")
            # Find the header row (first 20 rows) that contains "Specification"
            preview = pd.read_excel(self.fg_to_spec_path, engine="openpyxl", header=None, nrows=20)
            header_row = 0
            for i, row in preview.iterrows():
                if any(str(v).strip().lower() == "specification" for v in row.values):
                    header_row = i
                    break
            df = pd.read_excel(self.fg_to_spec_path, engine="openpyxl", header=header_row)

            # Find columns by name: ignores case, spaces and small name differences
            names = {c: str(c).strip().lower() for c in df.columns}
            spec_col = next((c for c, n in names.items() if n == "specification"), None) \
                or next((c for c, n in names.items() if "spec" in n), None)
            mat_col = next((c for c, n in names.items() if n == "material"), None) \
                or next((c for c, n in names.items() if "material" in n), None)
            if spec_col is None or mat_col is None:
                raise ValueError(
                    f"Specification/Material column not found in {os.path.basename(self.fg_to_spec_path)}. "
                    f"Columns found: {', '.join(map(str, df.columns))}"
                )
            logging.info(f"   Header row: {header_row + 1} | Spec column: '{spec_col}' | Material column: '{mat_col}'")

            # Same key format as FG Code in process_green_dot_file (12 digits, leading zeros)
            spec_keys = (df[spec_col].astype(str).str.strip()
                         .str.replace(r"\.0$", "", regex=True).str.zfill(12))
            materials = (df[mat_col].astype(str).str.strip()
                         .str.replace(r"\.0$", "", regex=True)
                         .replace({"nan": "Not Found", "None": "Not Found", "": "Not Found"})
                         .fillna("Not Found"))

            self.fg_dict = dict(zip(spec_keys, materials))
            
            logging.info(f"✅ Loaded {len(self.fg_dict)} FG mappings")
            
        except Exception as e:
            logging.exception("❌ Failed to load FG to Spec")
            self.fg_dict = {}
    
    def process_green_dot_file(self, file_path: str):
        """Process a single Green Dot file - FULLY OPTIMIZED"""
        try:
            filename = os.path.basename(file_path)
            logging.info(f"📄 Processing: {filename}")
            start_time = time.time()
            
            # Load file
            df = pd.read_excel(file_path, engine="openpyxl")

            # Remove exact duplicate rows
            dup_count = df.duplicated().sum()
            if dup_count:
                logging.warning(f"   Removed {dup_count:,} exact duplicate rows")
                df = df.drop_duplicates()

            # Auto-detect columns
            col_map = {col.lower().replace(" ", "").replace("_", ""): col for col in df.columns}
            
            fg_col = next((col_map[k] for k in col_map if "fg" in k and "code" in k), df.columns[0])
            cp_col = next((col_map[k] for k in col_map if "cp" in k and "code" in k), None)
            qty_col = next((col_map[k] for k in ["quantity", "cpquantity", "qty"] if k in col_map), None)
            bom_col = next((col_map[k] for k in col_map if "bom" in k and "base" in k), "Bom Base Quantity")
            site_col = next((col_map[k] for k in col_map if "site" in k or "production" in k), "FG Production Site")
            ec_col = next((col_map[k] for k in col_map if "ec" in k and "packaging" in k), "EC Packaging Level")
            
            if not all([fg_col, cp_col, qty_col]):
                logging.warning(f"Missing required columns in {filename}")
                return None
            
            # Melt material columns - detected by the " - " pattern, not by position
            material_cols = [c for c in df.columns if " - " in str(c)]
            fixed_cols = [c for c in df.columns if c not in material_cols]
            
            if not material_cols:
                logging.warning(f"No material columns found in {filename}")
                return None
            
            logging.info(f"   Fixed cols: {len(fixed_cols)} | Material cols: {len(material_cols)}")
            
            melted = df.melt(
                id_vars=fixed_cols,
                value_vars=material_cols,
                var_name="Constituent_Raw",
                value_name="Weight"
            )
            
            # Filter valid weights
            melted["Weight"] = pd.to_numeric(melted["Weight"], errors="coerce")
            melted = melted.dropna(subset=["Weight"])
            melted = melted[melted["Weight"] > 0].copy()
            
            if melted.empty:
                logging.warning(f"No valid data in {filename}")
                return None
            
            # Split constituent - VECTORIZED
            split_data = melted["Constituent_Raw"].str.split(" - ", n=1, expand=True)
            melted["Constituent"] = split_data[0].fillna(melted["Constituent_Raw"])
            melted["ConstituentCategory"] = split_data[1].fillna("")
            
            # Remove unwanted rows
            melted = melted[~((melted["Constituent"] == "Cartonboard") & 
                              (melted["ConstituentCategory"] == "Box board"))]
            
            # ===== VECTORIZED WeightQuantityFactor Calculation =====
            # Convert to numeric once
            melted["cp_clean"] = melted[cp_col].astype(str).str.strip()
            melted["site_clean"] = melted[site_col].fillna("").astype(str).str.strip().str.upper()
            melted["ec_clean"] = melted[ec_col].fillna("").astype(str).str.strip()
            if bom_col in melted.columns:
                melted["bom_num"] = pd.to_numeric(melted[bom_col], errors="coerce").fillna(1)
            else:
                logging.warning(f"   '{bom_col}' column not present - defaulting Bom Base Quantity to 1")
                melted["bom_num"] = 1.0
            melted["qty_num"] = pd.to_numeric(melted[qty_col], errors="coerce").fillna(0)
            
            # Set bom to 1 where it's 0 to avoid division errors
            melted.loc[melted["bom_num"] == 0, "bom_num"] = 1
            
            # Calculate base_factor using vectorized conditions
            base_factor = pd.Series(1.0, index=melted.index)
            
            # Condition 1: Special CP code
            mask1 = melted["cp_clean"] == "NA0D9025001"
            base_factor.loc[mask1] = 2
            
            # Condition 2: FG has a Consumer row with bom=1 and qty=1000
            # Applies to all CP codes of that FG (all packaging levels)
            special_row = (melted["ec_clean"] == "Consumer") & (melted["bom_num"] == 1) & (melted["qty_num"] == 1000)
            special_fgs = melted.loc[special_row, fg_col].unique()
            mask2 = (~mask1) & melted[fg_col].isin(special_fgs)
            base_factor.loc[mask2] = melted.loc[mask2, "qty_num"] / 1000
            
            # Condition 3: Consumer (general)
            mask3 = (~mask1) & (~mask2) & (melted["ec_clean"] == "Consumer")
            base_factor.loc[mask3] = 1
            
            # Condition 4: Mariscala Factory
            mask4 = (~mask1) & (~mask2) & (~mask3) & (melted["site_clean"] == "MARISCALA FACTORY (PACKAGING)")
            base_factor.loc[mask4] = melted.loc[mask4, "qty_num"]
            
            # Condition 5: Empty site
            mask5 = (~mask1) & (~mask2) & (~mask3) & (~mask4) & (melted["site_clean"] == "")
            base_factor.loc[mask5] = 1
            
            # Condition 6: Default calculation
            mask6 = (~mask1) & (~mask2) & (~mask3) & (~mask4) & (~mask5)
            base_factor.loc[mask6] = melted.loc[mask6, "qty_num"] / melted.loc[mask6, "bom_num"]
            
            melted["WeightQuantityFactor"] = base_factor * melted["Weight"]
            
            # Clean up temporary columns
            melted.drop(columns=["cp_clean", "site_clean", "ec_clean", "bom_num", "qty_num"], inplace=True)
            
            # Normalize CP Codes
            melted["CP Code"] = melted[cp_col].astype(str).str.strip().str.upper().str.lstrip("0")
            
            # Map recycled content
            melted["% Recycled Content"] = melted["CP Code"].map(self.cp_dict_percent).fillna(0)
            
            # Calculate Recycled WQF - VECTORIZED
            melted["Recycled WQF"] = melted["WeightQuantityFactor"] * melted["% Recycled Content"] / 100
            
            # ===== Calculate Recycled Weight in kg =====
            # Assuming WeightQuantityFactor is in grams, convert to kg by dividing by 1000
            melted["Recycled Weight (kg)"] = melted["Recycled WQF"] / 1000
            
            # Map Material - VECTORIZED
            fg_codes = melted[fg_col].astype(str).str.zfill(12).str.strip()
            melted["Material"] = fg_codes.map(self.fg_dict).fillna("Not Found")
            
            # Add source file
            melted["Source_File"] = filename
            
            elapsed = time.time() - start_time
            logging.info(f"✅ Processed {filename} ({len(melted):,} rows) in {elapsed:.2f}s")
            
            return melted
            
        except Exception as e:
            logging.exception(f"❌ Error processing {file_path}")
            return None
    
    def save_output(self, merged_df: pd.DataFrame):
        """Save final output with progress indicators"""
        
        # --- COMMON CLEAN FUNCTION: always take the second part after '|' if it exists ---
        def take_second_part(value):
            if pd.isna(value):
                return value
            
            parts = [p.strip() for p in str(value).split("|") if p.strip()]
            
            if not parts:
                return ""
            
            # If at least two parts, always return the SECOND one
            if len(parts) >= 2:
                return parts[1]
            
            # Otherwise return the only part
            return parts[0]
        
        # --- CLEANUP: House column ---
        if "House" in merged_df.columns:
            merged_df["House"] = merged_df["House"].apply(take_second_part)
        
        # --- CLEANUP: Brand column ---
        if "Brand" in merged_df.columns:
            merged_df["Brand"] = merged_df["Brand"].apply(take_second_part)
        # --- END CLEANUP ---
        
        final_cols = [
            "FG Code", "Material", "Code Title", "FG Type", "Bom Base Quantity", "EAN/UPC",
            "FG Production Site", "Division", "House", "Brand", "Part Family",
            "EC Packaging Level", "CP Code", "Quantity", "CP Title", "Gross Weight", "Net Weight",
            "Constituent", "ConstituentCategory", "Weight", "WeightQuantityFactor",
            "% Recycled Content", "Recycled WQF", "Recycled Weight (kg)", "Source_File"
        ]
        
        final_cols = [col for col in final_cols if col in merged_df.columns]

        # Keep any extra source columns (new report format), placed before "Constituent"
        extra_cols = [col for col in merged_df.columns
                      if col not in final_cols and col != "Constituent_Raw"]
        if extra_cols:
            logging.info(f"Extra columns kept: {', '.join(map(str, extra_cols))}")
            pos = final_cols.index("Constituent") if "Constituent" in final_cols else len(final_cols)
            final_cols = final_cols[:pos] + extra_cols + final_cols[pos:]

        merged_df = merged_df[final_cols]
        
        logging.info(f"💾 Saving {len(merged_df):,} rows to Excel...")
        save_start = time.time()
        
        merged_df.to_excel(self.output_path, index=False, engine='openpyxl')
        
        save_time = time.time() - save_start
        
        logging.info("="*60)
        logging.info(f"✅ SUCCESS! Output saved: {self.output_path}")
        logging.info(f"📊 Total rows: {len(merged_df):,}")
        logging.info(f"📁 Files processed: {merged_df['Source_File'].nunique()}")
        logging.info(f"⏱️ Save time: {save_time:.2f}s")
        logging.info("="*60)
    
    def run(self):
        """Main execution"""
        try:
            # File selection
            if not self.select_files():
                return
            
            # Setup logging
            self.setup_logging()
            
            # Load reference data
            self.load_environmental_report()
            self.load_fg_to_spec()
            
            if not self.cp_dict_percent or not self.fg_dict:
                logging.error("❌ Failed to load required reference data")
                return
            
            # Find and process files
            files = [os.path.join(self.folder, f) for f in os.listdir(self.folder)
                    if f.lower().startswith("green dot report") and f.lower().endswith(".xlsx")]
            
            if not files:
                logging.warning("⚠️ No Green Dot Report files found")
                return
            
            logging.info(f"Found {len(files)} files to process")
            
            all_results = []
            for file_path in files:
                result = self.process_green_dot_file(file_path)
                if result is not None:
                    all_results.append(result)
            
            if not all_results:
                logging.warning("⚠️ No files processed successfully")
                return
            
            # Merge
            logging.info(f"📦 Merging {len(all_results)} dataframes...")
            merged_df = pd.concat(all_results, ignore_index=True)
            
            # Save
            self.save_output(merged_df)
            
            # Success message
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo(
                "Complete",
                f"✅ Successfully processed {len(all_results)} files!\n\n"
                f"Total rows: {len(merged_df):,}\n\n"
                f"Output: {self.output_path}"
            )
            root.destroy()
            
        except Exception as e:
            logging.exception("❌ Fatal error")
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Error", f"Error: {str(e)}")
            root.destroy()


def main():
    """Entry point"""
    processor = GreenDotProcessor()
    processor.run()
    
    # Keep console window open
    print("\n" + "="*60)
    input("Press Enter to close this window...")


if __name__ == "__main__":
    main()
