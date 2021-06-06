# TGM-Tally-GSTR-Matching
Tally Data to GSTR Data Reconciliation Program.

### Description
* The program matches the Invoice from GST Portal to the Invoice from
Tally Data. It matches the GSTIN Number with their Invoice Number.
* The program reduces the manual work by almost **75% – 80%**.
* The program also creates a separate Excel File with contains the Reconciliation and **highlights** the data that has the exact match from GST Portal to Tally Data.
* The program also contains an **authentication** so that only a particular user can have access.
  * to add authentication edit .py file and replace 'ENTER_YOUR_PASSWORD_HERE' to 'YOUR_PASSWORD'


### WARNING ⚠
* **DO NOT** change the column names and excel sheet_name in TGM.xlsx (you can change the file name though)
* Run the program in python command line **ONLY**

### Disclaimer
* Data from the Excel sheets has been removed for **PRIVACY PURPOSE**
* An example of **Final Excel Sheet** has also been added to get an overview.
* Suggestion:
  * To hide the code from anyone seeing the password -> use **pyarmor** to obfuscate python scripts (https://pypi.org/project/pyarmor/)
