# Excel VBA Automation — Public Showcase (Demo)

> This repository showcases three Excel VBA automations using only fake data, mock flows, and simplified code skeletons.  
> It is designed for portfolio purposes and contains no company information.

What’s included
  - Monthly checklist backup (demo)
  - Mock flow for external system screen-capture (no real system calls)
  - Outlook email draft automation
What’s not included
  - No company names, codes, servers, or proprietary logic
  - No real data, no credentials, no integration secrets

Features

1. Monthly Checklist Backup (Demo
   - Creates a month-tagged copy (e.g., `Checklist_202501.xlsm`) in the same folder
   - Clears predefined input ranges for the new month
   - Safe, local-only behavior

2. External System Flow (Mock)
   - Simulates opening a transaction/report, waiting for user input, and “pasting” a screenshot
   - No actual system dependency (works on any machine)
   - Logs actions to a sheet as traceable steps

3. Outlook Email Draft
   - Drafts an email using values from a sheet (subject/to)
   - Attaches the current workbook
   - Uses `.Display` only (never `.Send` by default)

Process Diagram

> Conceptual only; no vendor-specific details.

```mermaid
flowchart TD
  A[User clicks Run Macro] --> B{Select Feature}
  B -->|Backup| C[Create YYYYMM copy]
  C --> D[Clear input ranges for next month]
  D --> E[Show success message]

  B -->|External Flow (Mock)| F[Log planned steps]
  F --> G[Prompt user for manual review]
  G --> H[Paste mock screenshot to Output sheet]

  B -->|Email Draft| I[Read Subject/To from Send_Email sheet]
  I --> J[Attach current workbook]
  J --> K[Display draft in Outlook]
