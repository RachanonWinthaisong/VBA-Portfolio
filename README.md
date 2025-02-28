# VBA Portfolio - Sample Excel Automation

## Overview
This repository contains sample VBA scripts designed to automate Excel tasks, focusing on data filtering and processing.

## Filter_Blanks_And_Lookup

### Description
This macro automates the following tasks in an Excel worksheet:

1. **Filter column CR for blank values**: Filters data in column CR to only show rows where the value is blank.
2. **Clear contents in column N**: Clears any existing data in column N for the filtered rows.
3. **Apply XLOOKUP formula**: Places an XLOOKUP formula in column N to fetch data from a report sheet.
4. **Convert formulas to static values**: After applying the XLOOKUP formula, the macro converts formulas in column N and BY to their resulting static values and removes any errors like `#N/A`.
5. **Filter column N for specific values**: Filters column N for rows with values "1" or blank and updates columns P and S accordingly.

### Usage
1. Open Excel and go to the `Developer` tab.
2. Click on `Visual Basic` to open the VBA editor.
3. Create a new module and copy-paste the provided script into the module.
4. Close the editor and return to Excel.
5. Run the macro by pressing `Alt + F8`, selecting `Filter_Blanks_And_Lookup`, and clicking `Run`.
   
### Notes:
- This macro assumes that your data follows a specific structure, with columns `CR`, `N`, `P`, `S`, and `BY` being utilized.
- The XLOOKUP function used references data from a sheet named `Report`—ensure that the `Report` sheet is correctly set up and populated.
- This script works best with Excel's macro functionality enabled.

📌 *This is a sample script for learning purposes. No actual company data is included.*

---

## Additional Features / Future Enhancements:
- Ability to handle dynamic ranges for columns
- Integration with multiple data sources for more advanced lookups
