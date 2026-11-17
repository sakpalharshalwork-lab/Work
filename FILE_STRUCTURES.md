# File structures (columns only, no data)

## Green Dot Report (SAP export, sheet "SAP Document Export"/"SAPUI5 Export")
Wide format: fixed columns + ~115 material columns named `Group - Material`.

- Old format fixed columns: FG Code, Code Title, FG Type, Bom Base Quantity, EAN/UPC, FG Production Site, Division, House, Brand, Part Family, EC Packaging Level, CP Code, Quantity, CP Title, Gross Weight, Net Weight
- New format adds: Owner, Unit Of Measure (PC/MTR/KG/G), Related MC, Spec Type, Spec Owner, Status (U050/U100/U200/U210/U230)
- Material column groups: Plastic, Paper, Metal, Glass, Cartonboard, Corrugated, Other, Natural, Fiber (e.g. `Plastic - PP`, `Metal - Aluminum`)
- Contains exact duplicate rows (removed by processor)

## Environmental_Report.xlsx
CP Code | Component | Constituent | Ref | Target | UOM | % Recycled Content | Instance | Comments | Certifications | Count
- CP Code can repeat (one row per constituent)

## FGtoSpec v2.xlsx
Formula GUID | Formula Item Number | Formula Item Type | Specification Key | S | pecification | Material | (blank) | Component Type | Quantity | Quantity | Quantity | Unit of Measure | Specification Type
- "Specification" is split by SAP: `S` = first character, `pecification` = rest

## Merged master (green_dot_all_merged_<timestamp>.xlsx)
FG Code | Material | Code Title | FG Type | Bom Base Quantity | EAN/UPC | FG Production Site | Division | House | Brand | Part Family | EC Packaging Level | CP Code | Quantity | CP Title | Gross Weight | Net Weight | [Owner | Unit Of Measure | Related MC | Spec Type | Spec Owner | Status] | Constituent | ConstituentCategory | Weight | WeightQuantityFactor | % Recycled Content | Recycled WQF | Recycled Weight (kg) | Source_File

## Lookup inputs
- Material lookup: one column `Material`
- EAN lookup: one column `EAN`

## Lookup output (RESULT_*.xlsx / RESULT_EAN_*.xlsx)
Material | EAN | Code Title | Gross Weight | Net Weight |
per level (Consumer, Customer, Transport): {Level}_Metal, _Plastic, _Cartonboard, _Glass, _Paper, _Corrugated, _Other, _Natural, _Metals_without_Aluminum, _Aluminum, _% Recycled |
% Recycled Content | Recycled Weight (kg)
