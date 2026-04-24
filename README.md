# Applicant Data Cleansing Application
A Python-based automation tool designed to sanitize and validate applicant spreadsheets using Regular Expressions (Regex) and automated error logging.

Brief Description: Takes an "applicants" spreadsheet that represents people that have applied to a local COMP program beforehand and cleans the data so only the valid applicants appear. 

# Key Features
    Regex Validation: Automatically validates Canadian postal codes and email formats to ensure data integrity.
    Automated Sanitization: Cleanses raw spreadsheet data (Excel/CSV) to prepare it for database ingestion.
    Error Reporting: Generates a log of invalid entries for manual review, preventing "bad data" from entering the system.

# How To Run It
    Clone the repo
    Install dependencies
    Put script in the same directory as the spreadsheet you want to cleanse
    Run main.py
