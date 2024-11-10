# Inventory Estimation and Expiry Management Script

This script processes and integrates pharmaceutical inventory data from various sources, calculates supply estimates, and predicts stock expiration based on both system and manual estimates. It reads inventory data from multiple files, consolidates them, and generates supply forecasts and summaries in CSV and HTML formats.

## Features

1. **Data Consolidation**: Merges data from five inventory sources: BMT pharmacy, drug store, hiwa pharmacy, in-patient pharmacy, and out-patient pharmacy.
2. **Supply and Expiry Calculations**: Calculates the supply of each item in months based on two estimates—system and manual—and identifies when items will expire.
3. **Automated Reporting**: Outputs a summary report with consolidated inventory data and supply calculations in `FINAL.csv` and `FINAL.html`.

## Requirements

- **Python 3.x**
- **Pandas**: For data manipulation and calculations.
- **NumPy**: For numerical operations.

## Installation

1. Install the required packages using pip:

   ```bash
   pip install pandas numpy
   ```

2. Ensure the following Excel files are available in the script’s directory:

   - `bmt_pharmacy.xls`
   - `drug_store.xls`
   - `estimation_work.xlsx`
   - `hiwa_pharmacy.xls`
   - `in_patient.xls`
   - `out_patient.xls`

## Usage

To run the script, use:

```bash
python inventory_estimation.py
```

### Inputs

The script reads inventory data and estimation files:

1. **Inventory Data**: The following Excel files, each representing an inventory source:
   - `bmt_pharmacy.xls`
   - `drug_store.xls`
   - `hiwa_pharmacy.xls`
   - `in_patient.xls`
   - `out_patient.xls`
2. **Estimation Data**: The file `estimation_work.xlsx`, containing average daily dispensing quantities and manual/system estimates for each item.

### Outputs

The script generates the following output files:

1. **`FINAL.csv`**: Consolidated CSV report of inventory data with supply and expiry calculations.
2. **`FINAL.html`**: HTML format of the final report.
3. **`edit_df_man.csv` and `edit_df_sys.csv`**: Intermediate data summarizing manual and system supply estimates, respectively.

## Functionality Overview

### 1. **Data Cleaning and Preparation**

   - The script reads data from each input Excel file, trims unnecessary columns, and adds a source column to each dataframe.
   - Rows from each source dataframe are appended to a consolidated dataframe for processing.

### 2. **Expiry and Supply Calculations**

   - **System and Manual Estimates**: Uses daily dispensing quantities to calculate the number of months each item will last (`supply in months sys` and `supply in months man`).
   - **Expiration Check**: For items with upcoming expiry dates, the script calculates if an item will expire before the stock is depleted.
   - **Monthly Calculations**: All calculations are normalized to monthly values.

### 3. **Output and Cleanup**

   - Saves the main consolidated report to `FINAL.csv` and `FINAL.html`.
   - Temporary intermediate files (`all_in_one.csv`, `edit_df_man.csv`, `edit_df_sys.csv`, and `edit_qty.csv`) are deleted after processing.

## Example Output Format

The final report, `FINAL.csv`, includes columns such as:

- **Item Name**: Name of the pharmaceutical item.
- **Quantity in Stock**: Total quantity available.
- **Supply in Months**: Estimated supply in months for both system and manual methods.
- **Will Expire**: Days remaining until expiry.

Each entry provides a full summary of inventory status, estimated supply duration, and expiration predictions.

## Troubleshooting

- **File Not Found Errors**: Ensure all required files are in the correct format and located in the script’s directory.
- **Calculation Issues**: Verify data types in the input files. Quantities and expiry dates must be correctly formatted.

## License

This script is provided under the MIT License.
