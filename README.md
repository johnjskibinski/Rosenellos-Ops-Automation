# Rosenello’s Operations Automation

Operational automation system for Rosenello’s Windows, Doors, Roofing & Siding.

Primary systems:

- LeadPerfection (CRM / financial system of record)
- Google Calendar (operations scheduling driver)
- Google Sheets Measure Packets
- CompanyCam (photo/project system)

---

## Current Modules

### Calendar Automation (syncMeasures)
- Cleans and normalizes calendar data
- Creates/updates Measure Packets
- Matches or creates CompanyCam projects (duplicate-proof)
- Adds CompanyCam links to installs/services
- Stores state in Calendar extendedProperties

### Measure Packet Template Config
- Hidden CONFIG sheet
- Tab mappings
- Print orientation rules
- LeadPerfection document type mapping

### Measure Packet Sidebar UI
- Export tabs to PDF
- Upload PDFs to LeadPerfection via API
- Upload Full Packet (all tabs individually)

---

## Security

Secrets are NEVER stored in GitHub.

All credentials are stored in:
Google Apps Script → Script Properties

---

## Owner

John Skibinski
Installation Manager – Rosenello’s
