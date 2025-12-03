# 🖥 Overview
Gdrive-To-Gsheet is an advanced Google Apps Script utility designed to scan any Google Drive folder and generate a structured, formatted, and analysis-ready Google Sheet output.

It supports files, folders, metadata, deep folder trees, multiple profiles, resume-based execution, and custom column preservation — all in a user-friendly, reliable workflow.

Perfect for teams and individuals performing audits, compliance checks, inventory management, migrations, or simply trying to understand the structure of a complex Drive hierarchy. 
[Download Link](https://github.com/TechPoov/GDrive-to-Gsheet/releases)  ![GitHub all releases](https://img.shields.io/github/downloads/TechPoov/GDrive-to-Gsheet/total.svg)

# Features
## 1. Multi-Profile Scan Engine
Set up multiple scan jobs in the Profiles sheet. Each profile runs independently with its own Name, Folder ID, Mode, and Depth limits.

## 2. Three Powerful Scan Modes
Files — scans only files
Folders — scans only folders with folder-level metrics
Both — full content + folder structure, ideal for audits

## 3. Time-Sliced Smart Worker
Large folder trees often time out in Google Apps Script.
This tool avoids that by using a resume-safe, slice-based worker system:

Handles tens of thousands of items
Never exceeds script limits
Automatically continues in the background
Robust against interruptions

## 4. Automatic Resume

If a scan is interrupted (quota hit, script time, browser close):
State is saved
Tool picks up exactly where it stopped
No duplicates, no resets, no lost time

## 5. Custom Column Preservation
If you add custom columns (Owner, ReviewedBy, Status, etc.):
They are not deleted
The tool merges them back into new scan results
Perfect for manual tagging, governance, or cleanup workflows

## 6. Clean, Professional Output Format

Each output sheet includes:
Hyperlinked file or folder names
Fixed-width columns for Path and File/Name
Auto-resized metadata columns
Hidden technical columns (ID)
Optional raw URL column (if enabled)

## 7. Profile-Level Navigation

Upon completion:
The ProfileName becomes a hyperlink
Clicking it jumps directly to the output sheet
Older output is archived as <ProfileName>_Old

## 8. Lightweight & Admin-Friendly
No external libraries.
No add-ons.
Runs inside your Google Sheet.
No installation for end users.

# Who Needs This Tool
## IT, Infrastructure, and Workspace Administrators
Use it for:
File/folder inventories
Drive structure understanding
Migration planning
Security and permission reviews
Audit documentation
Google Drive audits
Access cleanup
Folder structure analysis
Migration preparation (Google → OneDrive, etc.)
Compliance documentation

## Team Leads, Project Managers, Department Heads
Use it to:
View team folder contents at a glance
Track what is stored where
Identify cleanup opportunities
Monitor project artifacts
Often need:

A high-level inventory
File counts
Insights into team Drive usage
Duplicate folder structures
Recently modified content


## Compliance, Legal, Finance, and Audit Teams
Use it for:
Timestamp-based evidence
Structured metadata exports
Compliance documentation
Drive activity snapshotting
Document trails
Evidence gathering
Age analysis
Modification tracking
Ownership tagging

## Educators, Nonprofits, Freelancers, Students
Use it to:
Organize course materials
Map archival resources
Prepare project handoff documents
Track academic or volunteer file sets

## Personal and Power Users
Use it when your Drive becomes too large or disorganized.
Perfect for deep organization, cleanup, or visualization.

Ideal for:
Organizing huge personal Drives
Tracking old or misplaced files
Understanding shared Drive structures
Anyone who has ever felt lost inside Google Drive benefits from this tool.

# Use Cases
## Full Drive Inventory
Map an entire Drive hierarchy into a sortable, filterable sheet.

## Migration Preparation
Before migrating to SharePoint, OneDrive, Dropbox, etc., assess:
Folder depth
File counts
Last modified activity
Structure complexity


## Cleanup & Optimization
Identify:
Empty folders
Unused or old files
Rare file types
Redundant structures

## Compliance & Audit Reporting
Generate metadata needed for:
Financial audits
Security reviews
Legal documentation
IT governance

## Project Documentation
Snapshot a Drive folder for project closure or client handoff.

## Custom Tagging Workflows
Because custom columns persist, users can add:
Review status
Ownership
Priority
Risk classification

# How It Works 
For installation steps, configuration details, and usage instructions, refer to the QuickStart Guide included in the Downloads section.

# 6. Benefits
## Saves Time & Manual Effort
Eliminates the need to manually explore deep Drive structures.
## Audit-Ready Reports
Includes metadata, hyperlinks, timestamps, and structured rows suitable for reporting.
## Zero Training Required
Non-technical users can operate the tool from a single menu click.
## Works with Any Analysis Tool
Output is compatible with:
Google Sheets
Excel
Power BI
CSV workflows
Data Studio / Looker

## Team-Friendly
Since everything runs inside Google Sheets, teams use it without installing anything.

# Download Contents
## Gdrive-To-Gsheet — Copyable Google Sheet Template  
Link in QuickStart or User Manual -  Ready to duplicate and start scanning immediately.

## QuickStart Guide
A concise one-page introduction.

## User Manual
Full documentation with screenshots and workflows.

# What This Tool Does Not Do
Does not delete, modify, or move Drive files
Does not change file permissions
Does not perform real-time Drive monitoring
Does not replace admin-level audit logs
Does not analyze content inside files (only metadata)
This tool is strictly read-only and metadata-focused for safety.

# Requirements
A Google Workspace or personal Google account
Access to the Drive folders you want to scan
One-time authorization during first execution

# Version History
V1.0 - Initial base version. 
# License
Distributed under the MIT License.
# Support
For help, feature requests, or feedback:
📧 techpoov+GDrive-to-GSheet@gmail.com

🌐 https://techpoov.github.io

