# AI-Lead Generation System

[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-V8-4285F4?logo=google&logoColor=white)](https://developers.google.com/apps-script)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Status: Active](https://img.shields.io/badge/Status-Active-success.svg)](#)

*An automated, serverless B2B lead generation & outreach engine built entirely within Google Sheets and Google Apps Script.*

## Table of Contents
- [Problem Statement](#problem-statement)
- [Solution Overview](#solution-overview)
- [Key Features](#key-features)
- [High-Level Workflow](#high-level-workflow)
- [Business Impact](#business-impact)
- [Technologies Used](#technologies-used)
- [Installation Guide](#installation-guide)
- [Screenshots](#screenshots)
- [Future Improvements](#future-improvements)
- [Contact Information](#contact-information)

## Problem Statement
Finding, qualifying, and reaching out to local businesses is a highly manual, time-consuming process. Sales teams and agencies often spend hours clicking through Google Maps, inspecting websites, hunting for contact emails, and writing repetitive outreach messages, leading to low efficiency and generic, ineffective emails.

## Solution Overview
The AI-Lead Generation System automates the entire top-of-funnel sales process. Operating completely within Google Sheets, the system automatically discovers local businesses, evaluates their online presence, finds public contact information, qualifies high-potential leads, and drafts highly personalized, problem-first outreach emails directly into your Gmail Drafts folder.

## Key Features
- **Automated Business Discovery**: Automatically finds local businesses based on target industry and location.
- **Intelligent Website Analysis**: Evaluates the target business's online presence to understand their digital footprint.
- **Contact Discovery**: Safely extracts publicly available contact information from business websites.
- **Smart Lead Qualification**: Filters businesses based on strict criteria, ensuring only highly qualified leads are passed to the next stage.
- **Consultative Email Generation**: Drafts personalized, context-aware emails that focus on operational friction rather than generic sales pitches.
- **Seamless Gmail Integration**: Securely pushes generated emails directly to your Gmail Drafts folder for final review (never auto-sends).

## High-Level Workflow
**Google Maps**
➔ **Business Discovery**
➔ **Website Analysis**
➔ **Email Discovery**
➔ **Lead Qualification**
➔ **Personalized AI Email Generation**
➔ **Gmail Draft Creation**

## Business Impact
- **Who it helps**: Agencies, freelancers, and B2B sales teams targeting local businesses.
- **Problems it solves**: Eliminates the manual grunt work of lead generation, data entry, and writing repetitive emails.
- **Manual vs. Automated Workflow**: Reduces a process that typically takes hours of manual clicking and typing into a background process that runs automatically while you focus on high-value tasks.
- **Time Savings**: Saves an estimated 10-15 minutes of manual research and writing *per lead*.
- **Improved Outreach Efficiency**: Ensures every outreach attempt is highly personalized, factual, and consultative, significantly improving response rates compared to mass email blasts.

## Technologies Used
- **Google Apps Script (V8 Runtime)**
- **Google Sheets (Data Storage & UI)**
- **Google Places API (New)**
- **Gmail API**
- **Google Drive API**
- **Clasp (CLI Management)**

## Installation Guide

### Prerequisites
- Node.js (v16.x or newer)
- A Google account with access to Google Sheets and Gmail

### Setup Instructions
1. Install the Google Clasp CLI:
   ```bash
   npm install -g @google/clasp
   ```
2. Authenticate Clasp with your Google account:
   ```bash
   clasp login
   ```
3. Clone this repository:
   ```bash
   git clone <your-repo-url>
   cd AI-Lead-Generation-System
   ```
4. Create a new Google Sheet, open **Extensions > Apps Script**, and copy your Script ID from **Project Settings**. 
5. Update `.clasp.json` with your Script ID:
   ```json
   {
     "scriptId": "YOUR_APPS_SCRIPT_ID_HERE",
     "rootDir": "."
   }
   ```
6. Push the code to Google Apps Script:
   ```bash
   clasp push
   ```

### Required APIs
Ensure the following Google Cloud services are enabled in your Google Cloud Project:
- Google Places API (New) (Requires an active API Key)
- Google Sheets API
- Google Drive API
- Gmail API

### How to Run the Project
1. Open the Google Spreadsheet.
2. From the custom menu, click **Lead Generator > Initialize Workbook**.
3. In the generated `Settings` sheet, add your **Google Places API Key** and outreach details (Your Name, Company).
4. Run **Lead Generator > Generate Leads** to begin discovering and qualifying businesses!

## Screenshots

### Main Menu
*(Please insert screenshot of the custom Google Sheets menu here)*

### Raw Data Sheet
*(Please insert screenshot showing discovered businesses before qualification here)*

### Qualified Leads Sheet
*(Please insert screenshot showing qualified leads with their extracted information here)*

### Gmail Draft Example
*(Please insert screenshot of a personalized draft sitting in your Gmail Drafts folder here)*

### Overall Workflow
*(Please insert screenshot or diagram of the workflow in action here)*

## Future Improvements
- Multi-threaded lead processing
- Advanced analytics and reporting dashboards
- CRM integration capabilities
- Enhanced geographical targeting logic

## Contact Information
- **Name**: Harshika Gahlot
- **Email**: harshikagahlot01@gmail.com
- **LinkedIn**: [Your LinkedIn Profile URL]
- **GitHub**: [Your GitHub Profile URL]
