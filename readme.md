# README.md — Python Email → Excel → SQL Automated ETL Pipeline
## Overview

This project is a fully automated Python ETL pipeline designed to:

- Download Excel attachments from Outlook

- Process, clean, and transform data

- Extract and combine SQL queries

- Update a central Excel reporting template

- Generate daily submission files

- Maintain audit logs

- Orchestrate all steps via a master script

It replicates real-world NHS-style reporting automation, safely anonymised for demonstration.

## Key Features
### 1. Automated Email Attachment Extraction (email_downloader.py)

- Searches Outlook inbox using COM automation

- Filters by sender + subject

- Saves only .xlsx attachments

- Handles Outlook closed / error states

- Limits to recent N emails to improve speed

### 2. Excel Data Processing & Formula Recalculation (excel_updater.py)

- Loads workbook using openpyxl

- Inserts rows, styles columns, applies NamedStyles

- Calculates summary metrics

- Copies values between sheets

- Forces recalculation using Excel COM

- Saves refreshed file

### 3. SQL Data Extraction (database_queries.py)

- Connects to SQL Server via pyodbc

- Executes multi-part SQL query

- Parameterised date logic

- Returns metrics (admissions, discharges, occupancy)

- Integrates directly into Excel update pipeline

### 4. Bed Management & Escalation Update (bed_management_update.py)

- Combines Excel source files

- Sums cell ranges safely

- Loads metrics from SQL + Excel

- Builds output submission file

- Saves with automated timestamp

### 5. JSON-Based Change Logging (file_monitor.py)

- Tracks daily file counts

- Stores results in a JSON log

- Detects increases in received files

- Triggers downstream scripts conditionally

### 6. Task Scheduler-Compatible Master Script (main_script.py)

- Opens Outlook silently

- Executes all ETL steps

- Captures stdout/stderr

- Closes Outlook safely

- Designed for daily automation

### Technologies Used

- Python 3.x

- openpyxl

- pyodbc

- pandas

- comtypes (Outlook + Excel COM)

- glob/os/time/subprocess

- JSON logging
