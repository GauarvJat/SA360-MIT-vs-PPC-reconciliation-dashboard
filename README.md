# SA360 MIT vs PR (PPC) Reconciliation Utility

Automated reconciliation pipeline that validates **Publicis People Cloud (PPC)** spend data against **SA360 platform** reports and **Lumina (MIT)** data for Haleon media campaigns, with results visualized in Power BI.

## Overview

This project automates the end-to-end QA process of comparing media platform spend (SA360) against PPC-sourced data (Databricks SQL) to identify discrepancies, flag mismatches, and surface records missing from either source — replacing a manual Excel-based reconciliation process.

## Project Workflow

1. **Report Generation** – SA360 reports are created per advertiser with consistent column structures.
2. **Scheduled Delivery** – Reports are scheduled to a shared mailbox (`pgskreporting`), delivered at a fixed time (13 hours behind IST).
3. **Email Organization** – An Outlook rule files incoming reports into a dedicated folder.
4. **Attachment Extraction** – A VBA macro (or Power Automate) saves report attachments from Outlook to a local/BOX folder.
5. **Data Merge** – A Python script merges all SA360 CSV/Excel exports into a single dataset.
6. **ETL & Enrichment** – New columns (`Brand`, `Market_Code`, `Platform`, country names) are derived from advertiser naming conventions; pharma/vaccine advertisers are excluded.
7. **PPC Integration** – The PPC export (Databricks-sourced) is cleaned, standardized, and aligned to the same schema.
8. **Aggregation & Comparison** – Platform, PPC, and Lumina datasets are aggregated and merged on `Year`, `Platform`, `Market_Code`, and `Brand`.
9. **Variance Calculation** – Spend differences (absolute and %) are calculated between Lumina/PPC and Platform/PPC.
10. **Visualization** – The final dataset feeds a Power BI dashboard showing overall spend, market-level splits, mismatches, and records missing from either source.

<img width="207" height="848" alt="image" src="https://github.com/user-attachments/assets/30fc11e0-075e-46b9-a746-a9df487fe09d" />


## Repository Structure

```
├── README.md
├── requirements.txt
├── MIT_VS_PR_SA360_Utility.py      # Main ETL & reconciliation script
├── VBA_Code_Module1.bas            # Outlook attachment extraction macro
├── Documentation/
│   ├── Process_Flow.png            # End-to-end pipeline diagram
│   ├── Data_Mapping.xlsx           # Column mapping: Platform vs PPC vs Lumina
│   └── User_Guide.pdf              # Power BI dashboard sample/walkthrough
```

> Note: The full Power BI `.pbix` file is not included due to underlying client data. A sample report export (`User_Guide.pdf`) is provided instead to illustrate the dashboard structure and output.

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Update the hardcoded folder paths in `MIT_VS_PR_SA360_Utility.py` (`folder_path`, `input_folder`, `Market_code_file`, `Country_code_file`, output folders) to match your local environment.
3. Place SA360 platform exports (CSV/XLSX), the PPC export, and the Lumina export in the expected folders.
4. Run the script section by section (it is structured as Jupyter-style `# %%` cells).

## Data Sources

| Source | Description |
|---|---|
| SA360 Platform Reports | Per-advertiser CSV/XLSX exports from Search Ads 360 |
| PPC (Publicis People Cloud) | Spend export sourced via Databricks SQL |
| Lumina | MIT spend benchmark data (GBP) |
| Market/Country Reference Files | Local lookup files mapping market codes to country names |

## Outputs

- `platform_SA360_merged_file.xlsx` – Merged & enriched SA360 platform data
- `ppc_updated.xlsx`, `platform_updated.xlsx`, `lumina_updated.xlsx` – Standardized datasets
- `final_aggregated_df.xlsx` – Final reconciled dataset with variance metrics, used as the Power BI data source

## Known Issues / Open Items

- Power Automate "Get Emails (V3)" filter query configuration for the attachment-saving flow is still being finalized.
- Pharma and vaccine advertisers (identified via `VAX` in advertiser naming) are excluded from platform data per business rules.

## Tech Stack

- Python (pandas, numpy, openpyxl, charset-normalizer)
- VBA (Outlook automation)
- Power BI (visualization)
- Excel (reference/lookup files)
