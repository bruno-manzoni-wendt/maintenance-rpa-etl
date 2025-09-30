# Modular RPA to ETL Maitanence data

A modular RPA/ETL in Python to extract Maintenance data from a web app to visualize in a Qlik Sense Dashboard.

## STAR SUMMARY

**SITUATION:**
The Maintenance Department relied on a web app called “Checklist” for technicians to log Corrective, Preventive, Predictive and others maintenance activities. However, the tool lacked visualization and was not integrated into the company’s data warehouse. As a result, the manager had to manually export data and build reports in Excel, which was time-consuming.

**TASK:**
Automate data extraction and transformation from the web app and connect to a Qlik Sense dashboard.

**ACTION:**
Developed a modular Python-based RPA/ETL to extract and transform 9 tables of maintenance data from the web app and connect it into Qlik Sense. Designed KPIs and created a Technician Scoring System to evaluate performance.

**RESULTS:**
Enabled an automated, daily updated Qlik Sense dashboard that replaced manual reporting and improved data reliability. Delivered a solution more than a year and a half before official data warehouse integration, providing management with early visibility into maintenance performance and valuable insights.


## How It Works

### Main Orchestrator (call_maintenance_exports.py)
This script serves as the central controller that:
- Validates Authentication: Checks if the user is logged into the Checklist web app and automatically handles login if needed
- Orchestrates Exports: Sequentially executes 9 specialized export scripts, each handling a specific maintenance data type.
- Triggers Data Refresh: Automatically loads the extracted data into Qlik Sense after all exports complete

### Utility Module (checklist_utils.py)
A centralized library of reusable RPA functions that standardize the extraction process:
- Navigation: Opens specific URLs and handles page interactions
- Export Operations: Automates the multi-step export workflow (filtering, scrolling, clicking export buttons)
- Error Recovery: Monitors for export failures using pixel color detection and automatically restarts failed processes
- Progress Tracking: Monitors export completion status with periodic page refreshes to prevent timeouts
- File Management: Handles file downloads and validates successful saves

### Export Scripts
Each maintenance type has a dedicated script that extracts Raw Data using the utility module to navigate to the specific checklist page and processes the raw export into a structured format:
- Consolidates current and historical data
- Pivots checklist responses into normalized columns
- Separates datetime fields into date and time components
- Extracts and organizes image URLs into separate sheets
- Saves Output: Generates clean Excel files ready for Qlik Sense ingestion

## Key Features

- Autonomous Operation: Handles authentication automatically, eliminating manual intervention
- Error Resilience: Detects export failures and auto-restarts scripts to ensure data integrity
- Modular Design: Centralized utility functions enable easy maintenance and scalability
- Data Quality: Combines current exports with historical data and applies consistent transformations
- Progress Monitoring: Real-time console feedback with timestamps and completion status
- Multi-Source Integration: Processes 9 different maintenance data sources into a unified data pipeline


> **Disclaimer:** This repository contains scripts that have been modified to replace sensitive information with example placeholders for privacy and portfolio purposes.
