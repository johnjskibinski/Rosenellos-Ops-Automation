# Architecture Overview

## Systems

- LeadPerfection (CRM / job finance system of record)
- Google Calendar (operations driver)
- Google Sheets (measure packet / print docs)
- CompanyCam (photo + project tracking)

## Automation Layers

### 1. Calendar Automation (syncMeasures)
- Cleans measure events
- Creates measure packets
- Links CompanyCam
- Adds CompanyCam to installs/services
- Stores state in event extendedProperties

### 2. Measure Packet Template Config
- Hidden CONFIG sheet
- Tab mappings
- Print orientation rules
- LP doc type mapping

### 3. Measure Packet Sidebar UI
- Export tab to PDF
- Upload to LP API
- Upload full packet (5 tabs individually)

## Security Model

- No secrets stored in GitHub
- No secrets stored in Sheets
- Secrets stored in Apps Script Script Properties
