# LeadPerfection API Endpoints

## Authentication

- POST /token

---

## Job Lookup / Financial Data

- POST SalesApi/GetProspectJobID
- POST SalesApi/GetProspectJobID2
- POST SalesApi/GetSalesJobDetail

---

## Document Upload

- POST UploadDocument

---

## Installer / Scheduling (Future Expansion)

- POST GetInstallerApptCal
- POST GetInstallerApptDetail
- POST GetInstallerJobDetail
- POST AddInstallerJobNotes

---

## Sales / Job Data (Future Expansion)

- POST GetSalesApptCal
- POST GetSalesApptDetail
- POST GetSalesJobDetail
- POST UpdateSalesJobDetail
- POST AddSalesJobCost
- POST UpdateSalesJobCost

---

## Notes / Files / Media (Future Expansion)

- POST AddNotes
- POST AddJobImages
- POST GetJobImages
- POST DownloadFiles

---

## Customers / Leads (Future Expansion)

- POST GetCustomers
- POST GetCustomers2
- POST GetLead
- POST GetLeadData
- POST GetProspectData

---

## Notes

- API authentication uses LP API user credentials (username/password + appkey).
- Uploads use API endpoints — not the web upload endpoint.
- Document uploads map to specific LP Document Type IDs.
- Job matching logic prefers phone, then confirms via address + name.
