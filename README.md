# Excel Sheet Automation Script

This script processes an Excel file to generate a new sheet with filtered and formatted data for new hires.

## Requirements

Install the necessary packages:

```bash
pip install pandas openpyxl
```

## Files to Setup

1. **input_file**: Your source Excel file (e.g., `OfferReport.xlsx`).
2. **output_file**: The output Excel file where the processed data will be saved (e.g., `NewHire_IT_Checklist.xlsx`).
3. **config.py**: Set the following values:
   ```python
   input_file = 'OfferReport.xlsx'  # Input file
   output_file = 'NewHire_IT_Checklist.xlsx'  # Output file
   start_date = "09/03/2024"  # The date to filter data by
   ```

## How to Use

1. **Prepare your input file** Download a copy of the FTE Sheet and upload it to VScode: 
   - `Candidate`, `Personal Email`, `Better Email`, `Job`, `Department`, `Start Date`, `TimeZone`.

3. **Set the start date** in `config.py` to filter data (format: `MM/DD/YYYY`).

4. **Run the script**: Execute `main.py`. It will create a new sheet in the output file based on the start date.

5. **Check the output**: The new sheet will be saved in the `output_file` and will be opened automatically.

