# Expense Claim System
Use [Google Form](https://www.google.com/forms/about/) to collect expense claims and then extract data via [Google Sheets API](https://developers.google.com/sheets/api) and [Google Drive API](https://developers.google.com/drive) to generate reports

## Features
* Combine separated receipts and evidences of exchange rate images as a single PDF respectively for each claim
* Automatic exchange rate conversion
* Automatic approval request process
* Update data on a daily basis
* Provide overall summary of all claims

## Directory Structure
    .
    ├── configFile                                  # Contains all the configuration files required in the scripts
    │   ├── converted_header.json                   # Column letters in Google Spreadsheet, which indicate converted HKD amount
    │   ├── receipt_header.json                     # Column letters in Google Spreadsheet, which indicate receipt image URL in Google Drive
    │   ├── evidence_header.json                    # Column letters in Google Spreadsheet, which indicate evidence of exchange rate image URL in Google Drive
    │   ├── email_info.json                         # Email info of the sender
    │   └── requirements.txt                        # Required Python packages
    ├── Fonts                                       # Set font style for text in receipt and evidence PDF
    ├── Template                                    # Place expense claim CSV template to create new claims
    ├── output                                      # Contains all the output files (automatically generated)
    │   ├── expense claim form                      # Store generated CSV and PDF claim forms
    │   ├── receipts                                # Download receipt images and create combined receipt PDF for each claim
    │   │   ├── image
    │   │   └── PDF
    │   ├── evidences                               # Download evidence of exchange rate images and create combined evidence PDF for each claim
    │   │   ├── image
    │   │   └── PDF
    │   └── summary                                 # Store CSV summary of claims for each month and each audit period
    └── Response_Extraction.py                      # Main Python script

## Solution Architecture
![Solution Architecture Illustration](https://github.com/andy2167565/Expense-Claim-System/blob/335ed46701aa741a23c5adb1a4e5fd4a49ab5344/approach_1.2.png)
1. Use Google Form to collect applicants’ responses, receipts and evidences of exchange rate in form owner’s Google Drive. 
      *	The responses are stored in Google Spreadsheet for extraction.
      *	The receipts and evidences of exchange rate are placed in a folder and separated by item number respectively.
2. Use Python to retrieve data automatically.
      *	Extract responses from Google Spreadsheet via Google Sheets API, and then generate each expense claim form as Excel and PDF with required template format.
      *	Download related receipts and evidences of exchange rate via Google Drive API at the same time.
3. Use Python to automatically send the PDF claim form to the applicants’ supervisors for approval.
4. Use Python to extract responses that finished the claim process and save responses as summary files.

## Logic Flow
1.	Enable [Google Sheets API](https://developers.google.com/sheets/api/quickstart/python) and [Google Drive API](https://developers.google.com/drive/api/v3/quickstart/python) to get token and credential
2.	Execute ```Response_Extraction.py```
4.	Connect to Google Spreadsheet
5.	Capture and download all claim data
6.	Save data as CSV, PDF and image files

## How to Execute
### Install packages
```
pip install -r requirements.txt
```

### Run the script
```
python Response_Extraction.py
```
***
Copyright © 2020 [Andy Lin](https://github.com/andy2167565). All rights reserved.
