# Daily Financial Breaks Report Automation

Python automation script designed to generate a daily End-of-Day (EOD) financial control report by consolidating reconciliation, cash balance, and mark-to-market (MTM) break information from multiple sources.

This project demonstrates workflow automation in a financial operations environment, reducing manual reporting and ensuring consistent daily monitoring of operational breaks.

---

## Overview

In many financial institutions, control teams must manually compile daily reports combining information from several systems and internal communications.

This script automates that process by:

1. Retrieving relevant reports from predefined file locations
2. Extracting key break information from incoming emails
3. Compiling the information into a structured EOD summary
4. Automatically generating a formatted email report
5. Saving the final report for documentation purposes

The goal is to reduce repetitive manual work and standardize the reporting workflow.

---

## Key Features

- Automated date handling (including weekend adjustments)
- Integration with Microsoft Outlook via `win32com`
- Automated retrieval of relevant email reports
- Parsing of report content from email bodies
- Automatic attachment of supporting files
- Generation of a structured HTML email report
- Automatic saving of the generated report

---

## Example Workflow

Typical daily process automated by the script:

1. Retrieve daily financial control reports
2. Extract exception information from incoming operational emails
3. Consolidate key break indicators:
   - MTM breaks
   - Cashpool and ctpy balance breaks
   - Reconciliation breaks
4. Generate a standardized end-of-day control report
5. Prepare and display the email for final review before sending
6. Save the report locally for record keeping

---

## Technologies Used

- Python
- pandas (data processing in related scripts)
- win32com (Outlook automation)
- datetime (automated date handling)
- os / pathlib (file management)

---

## Project Structure
