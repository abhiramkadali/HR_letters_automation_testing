# HR Letter Generator — Full Documentation

**Tool:** `hr_admin.html` + `hr_user.html`
**Version:** 6.0
**Prepared by:** HR Employee Services
**Last Updated:** June 2026

---

## Table of Contents

1. [What This Tool Does](#1-what-this-tool-does)
2. [Two-App Architecture](#2-two-app-architecture)
3. [How the Data Flow Works](#3-how-the-data-flow-works)
4. [Using the Admin App — hr_admin.html](#4-using-the-admin-app--hr_adminhtml)
5. [Using the User App — hr_user.html](#5-using-the-user-app--hr_userhtml)
   - 5.1 [Connecting to SharePoint or GitHub](#51-connecting-to-sharepoint-or-github)
   - 5.2 [Uploading the ServiceNow File](#52-uploading-the-servicenow-file)
   - 5.3 [Generating a Letter](#53-generating-a-letter)
   - 5.4 [Downloading the Letter](#54-downloading-the-letter)
6. [Excel File Formats — Required Columns](#6-excel-file-formats--required-columns)
   - 6.1 [HRIS File — Employee Master Data](#61-hris-file--employee-master-data)
   - 6.2 [ServiceNow File — Letter Requests](#62-servicenow-file--letter-requests)
   - 6.3 [Valid Location Codes](#63-valid-location-codes)
   - 6.4 [Sample Files](#64-sample-files)
7. [Letter Types — All 11 Templates](#7-letter-types--all-11-templates)
8. [Letter Features](#8-letter-features)
9. [How to Make Changes to the Code](#9-how-to-make-changes-to-the-code)
   - 9.1 [Update an Office Address or Phone Number](#91-update-an-office-address-or-phone-number)
   - 9.2 [Add a New Office Location](#92-add-a-new-office-location)
   - 9.3 [Edit Wording in a Letter Template](#93-edit-wording-in-a-letter-template)
   - 9.4 [Add a New Letter Type](#94-add-a-new-letter-type)
   - 9.5 [Change the Salary Split Percentages](#95-change-the-salary-split-percentages)
   - 9.6 [Add or Edit Addressee Quick-Pick Options](#96-add-or-edit-addressee-quick-pick-options)
   - 9.7 [Change the Logo Size](#97-change-the-logo-size)
   - 9.8 [Change the Signatory Title](#98-change-the-signatory-title)
   - 9.9 [Change Brand Colours](#99-change-brand-colours)
   - 9.10 [Update ServiceNow Column Names](#910-update-servicenow-column-names)
   - 9.11 [Add, Edit, or Remove Authorised Signatories](#911-add-edit-or-remove-authorised-signatories)
   - 9.12 [Update the Verification Email Address](#912-update-the-verification-email-address)
10. [Code Structure Map](#10-code-structure-map)
11. [Troubleshooting](#11-troubleshooting)
12. [Frequently Asked Questions](#12-frequently-asked-questions)

---

## 1. What This Tool Does

The HR Letter Generator is a self-contained web application split into two HTML files. No server installation, no database, and no coding knowledge is required to operate it day-to-day.

The system generates formal HR letters and certificates for employees across all UAE offices. It handles 7 office locations and 11 letter types — 77 combinations — each with the correct office letterhead, employee data, and formatting applied automatically.

**Key capabilities:**

- Admin uploads and validates the full HRIS employee master data periodically
- HRIS data is published to a shared location (SharePoint, OneDrive, or GitHub) and loaded automatically by the user app
- HR officers upload only the daily ServiceNow report — the employee list shows only today's requestors
- Letter type, addressee, purpose, and travel dates auto-fill from the ServiceNow ticket
- Every letter carries a confidential banner, the request number, closing statement, and verified footer note
- Letters can be downloaded as PDF (directly or after editing) or as an editable Word document
- Downloaded files are named automatically: `RITM_EmployeeName_LetterType`

---

## 2. Two-App Architecture

| App | File | Who uses it | When |
|---|---|---|---|
| **Admin App** | `hr_admin.html` | IT / HR Admin | Periodically — when HRIS data is refreshed |
| **User App** | `hr_user.html` | HR Officers | Daily — to process letter requests |

Both files are hosted on SharePoint (or a shared network location) so everyone always has the latest version.

```
hr_admin.html                   Shared Location (SharePoint / GitHub)
─────────────────               ──────────────────────────────────────
Admin uploads HRIS.xlsx  ──▶    hris_data.json   ◀── hr_user.html
Validates employee data          hr_admin.html        fetches on open
Publishes hris_data.json         hr_user.html
                                 HRIS_Sample.xlsx
                                 ServiceNow_Sample.xlsx
```

---

## 3. How the Data Flow Works

```
HRIS Export (all employees)       ServiceNow Report (today's requests)
         │                                      │
         │  Admin validates + publishes         │  User uploads daily
         ▼                                      ▼
   hris_data.json               Auto-joined on Email Address
   (on SharePoint)                              │
         │                                      │
         └──────────────┬─────────────────────┘
                        ▼
            Merged employee list
            (only requestors shown)
                        │
             ┌──────────┴──────────┐
             │  Auto-filled form   │
             │  Letter type        │
             │  Addressee          │
             │  Purpose            │
             │  Travel dates       │
             └──────────┬──────────┘
                        ▼
              Generate → Edit → PDF / Word
```

**The join key is the employee's work email address.** Every ServiceNow row is matched to its HRIS record by email. Unmatched rows appear in red so HR can investigate.

**ServiceNow overrides HRIS** for Passport Number and Nationality — useful when an employee has submitted a renewed passport with their request.

---

## 4. Using the Admin App — hr_admin.html

### Step 1 — Set the SharePoint folder URL (once only)

Enter the full `https://` URL of the SharePoint folder where `hris_data.json` will be stored. Click **Save URL** — this is remembered in your browser.

### Step 2 — Upload the HRIS Export

Click the upload zone or drag and drop your HRIS `.xlsx` export. The app:

- Reads all rows and validates every employee record
- Flags missing names, emails, passport numbers, unknown location codes, and missing salaries
- Shows a full employee table with a status badge per row (**OK** or **⚠ issues**)
- Displays summary statistics: total employees, offices, departments, rows with issues

Use **Show issues only** to filter to problem rows. Fix issues in the source file and re-upload.

If column headers in your export differ from expected names, a **column mapping panel** appears — use the dropdowns to match your columns, then click **Apply Mapping & Load Data**.

### Step 3 — Publish to SharePoint

Click **📤 Publish to SharePoint**. The app:

1. Downloads `hris_data.json` to your computer
2. Shows a direct link to your SharePoint folder
3. Provides step-by-step instructions to upload the file

Upload `hris_data.json` to the SharePoint folder. The user app will load it automatically on its next open.

> **How often to publish:** Whenever the HRIS export changes — typically after new hires, leavers, salary changes, or role changes. The user app always loads the most recently published version.

---

## 5. Using the User App — hr_user.html

### 5.1 Connecting to SharePoint or GitHub

The first time the user app is opened:

1. Click **change** next to the SharePoint URL field in the HRIS section
2. Paste the SharePoint folder URL (same URL the admin used)
3. Click **Save**

The app immediately fetches `hris_data.json` from that URL and shows a green status with the employee count and publish timestamp. This URL is saved in the browser — you only need to set it once.

**If the fetch fails** (wrong URL, network issue, or file not yet published): a red status appears and a fallback HRIS upload zone is revealed. You can upload `HRIS_Sample.xlsx` or your HRIS export directly as a workaround.

**If no URL is configured:** the app uses built-in sample data (25 employees) — sufficient for testing but not for live use.

### 5.2 Uploading the ServiceNow File

Upload today's ServiceNow report into the **ServiceNow** upload zone. The app:

- Joins every request row to its HRIS record via email address
- Replaces the employee list with only today's requestors
- Shows a merge banner: ✅ green if all matched, ⚠️ amber if any unmatched
- Collapses the upload panel into a compact summary line

Once loaded, the upload panel collapses and the **Today's Requests** list expands to show all request cards.

### 5.3 Generating a Letter

| Step | Action |
|------|--------|
| 1 | Click a request card in the **Today's Requests** list |
| 2 | Letter type, purpose, addressee, country, and travel dates auto-fill from the ServiceNow ticket |
| 3 | Adjust any auto-filled fields if needed |
| 4 | Select the **Letter Date** (defaults to today) |
| 5 | Select the **Authorised Signatory** from the dropdown |
| 6 | Click **Generate Letter →** |
| 7 | Review the letter preview on the right |

### 5.4 Downloading the Letter

Three download options are available from the preview toolbar after generating a letter:

| Button | What it does |
|---|---|
| **Print / Save PDF** | Opens a clean print window — set Destination to "Save as PDF" in the print dialog |
| **✏️ Edit & Download Word** | Opens an editable version of the letter — click any paragraph to amend, then click **⬇ Download Word (.doc)** |
| **⬇ Download PDF** (in edit window) | Saves a PDF of the edited letter directly from the edit window |

All downloaded files are named automatically: `RITM0012301_Mohammed_Al_Rashidi_Salary_Certificate.pdf`

> **PDF tip:** In Chrome, enable **Background graphics** in the print dialog so the dark table header rows print correctly.

---

## 6. Excel File Formats — Required Columns

### 6.1 HRIS File — Employee Master Data

Column order does not matter — only the header names matter. The app uses fuzzy matching, so minor variations (e.g. "Emp ID" instead of "Employee ID") are caught automatically or prompted via the mapping panel.

| Column Header | Required | Description | Example |
|---|---|---|---|
| `Employee ID` | Recommended | Unique employee code | EMP001 |
| `Name` | **Yes** | Full name as it appears on the letter | Mohammed Al Rashidi |
| `Email` | **Yes** | Work email — **join key** for matching ServiceNow rows | m.rashidi@company.com |
| `Designation` | Yes | Job title | Senior Financial Analyst |
| `Department` | Yes | Department name | Finance |
| `Location Code` | Yes | Must match an office code (see Section 6.3) | DXB |
| `Join Date` | Yes | Date of joining | 15 March 2019 |
| `Nationality` | Yes | Nationality | Emirati |
| `Passport Number` | Yes | Passport number | A12345678 |
| `Basic Salary` | Yes | Basic salary component | AED 12,025 |
| `Housing Allowance` | Yes | Housing allowance | AED 4,625 |
| `Transport Allowance` | Yes | Transport allowance | AED 1,850 |
| `Total Salary` | Yes | Total monthly salary | AED 18,500 |
| `Bank` | Yes | Employee's bank name | Emirates NBD |
| `IBAN` | Yes | IBAN number | AE07 0331 2345 6789 0123 456 |

> **Salary tip:** If your HRIS only exports a Total Salary with no breakdown, the app auto-calculates Basic (65%), Housing (25%), and Transport (10%) as a fallback. See Section 9.5 to change these percentages.

### 6.2 ServiceNow File — Letter Requests

Only employees present in this file will appear in the Today's Requests list.

| Column Header | Required | Description | Example |
|---|---|---|---|
| `Email` | **Yes** | Work email — **join key** — must match HRIS | m.rashidi@company.com |
| `Request Number` | Recommended | ServiceNow RITM number — used in filenames | RITM0012301 |
| `Letter Type` | Recommended | Requested letter type — matched to templates | Salary Certificate |
| `Addressed To` | Optional | Recipient name / role | The Branch Manager |
| `Country` | Optional | Recipient country — appended to addressee line | UAE |
| `Purpose` | Optional | Reason stated in the ticket | bank account upgrade |
| `Travel Date From` | Optional | Travel start date — required for NOC tourist/business | 10 July 2026 |
| `Travel Date To` | Optional | Travel end date | 25 July 2026 |
| `Passport Number` | Optional | Overrides HRIS passport if provided | A12345678 |
| `Nationality` | Optional | Overrides HRIS nationality if provided | Emirati |
| `Additional Notes` | Optional | Free-text notes from the requestor | Please issue urgently |

> **Letter Type matching:** The "Letter Type" value is matched against template names using fuzzy matching. For reliable auto-selection use names close to: `Salary Certificate`, `Salary Transfer Letter`, `Employment Certificate`, `No Objection Certificate`, `NOC Tourist Visa`, `NOC Business Visa`, `Experience Letter`, `Visa Support Letter`, `Bank Account Opening`, `Address Proof Letter`, `Reference Letter`.

### 6.3 Valid Location Codes

The `Location Code` in the HRIS file must exactly match one of the codes below (case-sensitive).

| Code | Office |
|------|--------|
| DXB | Dubai – Head Office |
| AUH | Abu Dhabi |
| SHJ | Sharjah |
| AJM | Ajman |
| RAK | Ras Al Khaimah |
| UAQ | Umm Al Quwain |
| FUJ | Fujairah |

### 6.4 Sample Files

| File | Contents |
|---|---|
| `HRIS_Sample.xlsx` | 25 sample employees across all 7 UAE offices — includes intentional test cases (missing passports, unknown location code, salary-only row) |
| `ServiceNow_Sample.xlsx` | 15 sample requests covering all 11 letter types — includes duplicate requests, passport overrides, travel dates, and one unmatched email |

Both files have a **ReadMe** tab explaining the test cases and column notes.

---

## 7. Letter Types — All 11 Templates

| # | ID | Name | Key data used |
|---|---|---|---|
| 1 | `salary_cert` | Salary Certificate | Full salary breakdown, passport, nationality |
| 2 | `salary_transfer` | Salary Transfer Letter | Salary breakdown, bank, IBAN |
| 3 | `employment_cert` | Employment Certificate | Designation, department, salary, join date |
| 4 | `noc` | No Objection Certificate | Designation, department, purpose |
| 5 | `noc_tourist` | NOC — Tourist Visa | Travel dates from/to, destination country, work location |
| 6 | `noc_business` | NOC — Business Visa | Travel dates from/to, destination country |
| 7 | `experience` | Experience Letter | Designation, join date, department |
| 8 | `visa_support` | Visa Support Letter | Salary breakdown, passport, travel purpose |
| 9 | `bank_account` | Bank Account Opening | Salary breakdown, bank, passport |
| 10 | `address_proof` | Address Proof Letter | Residential address |
| 11 | `reference` | Reference Letter | Designation, join date, department |

**Templates 5 and 6 (NOC Tourist and NOC Business)** pull travel dates and destination country directly from the ServiceNow report. If those fields are populated in ServiceNow, they auto-fill when the request card is clicked.

---

## 8. Letter Features

Every generated letter includes the following automatically — no manual input required:

| Feature | Detail |
|---|---|
| **Confidential banner** | "🔒 PRIVATE & CONFIDENTIAL" strip at the top of every letter |
| **Office letterhead** | Correct office name, address, phone, and email based on employee's Location Code |
| **Request Number** | ServiceNow RITM number displayed in the letter header (e.g. RITM0012301) |
| **Closing statement** | "This certification is issued upon the Employee's request, without any liability on the part of the Firm. The Firm confirms that the information provided in this certificate is accurate as of the date of issuance and may be subject to amendment in the future." — appears as the final paragraph of every letter body |
| **Authorised signatory** | Selected signatory's name and designation (no office or location line) |
| **Footer note** | "This document is electronically generated and e-signed following authorization by the signatory. For further inquiries or verifications, please contact the firm's verifications team at empservices@test.com" |
| **Font** | Aptos (Calibri fallback) — consistent with Microsoft Office 2024+ standard |
| **Logo** | Deloitte wordmark at 192×37px (proportional to 320×61px reference) |

---

## 9. How to Make Changes to the Code

All changes are made by opening the relevant `.html` file in **Notepad** (right-click → Open with → Notepad) and editing the relevant section. Use **Ctrl + F** to search for the text shown below.

Changes to `hr_user.html` should be re-uploaded to SharePoint after editing so all users get the updated version.

---

### 9.1 Update an Office Address or Phone Number

**Search in:** `hr_user.html` (and `hr_admin.html` for validation purposes)
**Search for:** `SECTION 1: OFFICE LOCATIONS`

```javascript
{ id:"DXB", name:"Dubai – Head Office", address:"Level 12, Dubai World Trade Centre...", phone:"+971 4 123 4567", email:"hr.dubai@company.com" },
```

Find the office by its `id` and update the relevant field. Save the file.

---

### 9.2 Add a New Office Location

**Search for:** `SECTION 1: OFFICE LOCATIONS`

Copy any existing line and paste it below the last entry:

```javascript
{ id:"XYZ", name:"New Office Name", address:"Full address", phone:"+971 X XXX XXXX", email:"hr.xyz@company.com" },
```

> The `id` must exactly match the `Location Code` value in the HRIS file for employees at that office.

---

### 9.3 Edit Wording in a Letter Template

**Search for:** `SECTION 5: LETTER TEMPLATES`

Each template is labelled, e.g. `// ── TEMPLATE 1: Salary Certificate`. Find the template and change the paragraph text. Do not remove placeholders.

**Available placeholders:**

| Placeholder | What it inserts |
|---|---|
| `${emp.name}` | Employee full name |
| `${emp.designation}` | Job title |
| `${emp.department}` | Department |
| `${emp.joinDate}` | Date of joining |
| `${emp.nationality}` | Nationality (ServiceNow override if provided) |
| `${emp.passportNo}` | Passport number (ServiceNow override if provided) |
| `${emp.salary}` | Total monthly salary |
| `${emp.basic}` | Basic salary |
| `${emp.housing}` | Housing allowance |
| `${emp.transport}` | Transport allowance |
| `${emp.bank}` | Bank name |
| `${emp.iban}` | IBAN number |
| `${emp.address}` | Residential address |
| `${loc.name}` | Office name |
| `${loc.email}` | Office HR email |
| `${purpose}` | Purpose (auto-filled from ServiceNow) |
| `${addressee}` | Addressee + country (auto-filled from ServiceNow) |
| `${emp._snReq?.snTravelFrom}` | Travel start date (from ServiceNow) |
| `${emp._snReq?.snTravelTo}` | Travel end date (from ServiceNow) |
| `${emp._snReq?.snCountry}` | Destination country (from ServiceNow) |

---

### 9.4 Add a New Letter Type

**Step 1 — Register the letter type:**
Search for `SECTION 2: LETTER_TYPES` and add a new entry:

```javascript
{ id:"your_letter_id", name:"Your Letter Name", description:"Short description" },
```

**Step 2 — Add the template body:**
Search for `SECTION 5: LETTER TEMPLATES` and add before the closing `};`:

```javascript
your_letter_id: (emp, loc, purpose, addressee) => `
  <p class="salutation">To: ${addressee},</p>
  <p class="subject-line">SUBJECT: YOUR SUBJECT HERE</p>
  <p>Body text using ${emp.name}, ${emp.designation}, etc.</p>
`,
```

> The `id` in both steps must match exactly.

---

### 9.5 Change the Salary Split Percentages

**Search for:** `TEMPLATE 1: Salary Certificate`

```javascript
const basicAmt = emp.basic || "AED " + Math.round(total * 0.65).toLocaleString();
const housAmt  = emp.housing || "AED " + Math.round(total * 0.25).toLocaleString();
const tranAmt  = emp.transport || "AED " + Math.round(total * 0.10).toLocaleString();
```

Change `0.65`, `0.25`, `0.10` to your desired percentages. They must add up to 1.0.

---

### 9.6 Add or Edit Addressee Quick-Pick Options

**Search for:** `SECTION 6: ADDRESSEE PRESETS`

```javascript
const ADDRESSEE_PRESETS = [
  { label: "Bank",    value: "The Branch Manager" },
  { label: "Embassy", value: "The Consulate General / Embassy" },
  ...
];
```

- `label` → button text shown in the sidebar
- `value` → text printed on the letter

---

### 9.7 Change the Logo Size

**Search for:** `width="192" height="37"`

Change the `width` and `height` values. To maintain the correct aspect ratio, keep the ratio at approximately **5.25:1** (width:height). Examples:

| Width | Height | Scale |
|---|---|---|
| 160 | 30 | ~50% of reference |
| 192 | 37 | 60% of reference (current) |
| 224 | 43 | ~70% of reference |
| 256 | 49 | 80% of reference |

---

### 9.8 Change the Signatory Title

The signatory is selected from a dropdown at generation time — see Section 9.11 to update the list. The label "Authorised Signatory" shown as fallback when none is selected can be changed by searching for `Authorised Signatory` in the code.

---

### 9.9 Change Brand Colours

**Search and replace in Notepad** (`Ctrl + H`):

| Current value | What it controls |
|---|---|
| `#0f172a` | Dark navy — sidebar background, borders, table header row |
| `#2563eb` | Blue — Generate button, active selections |
| `#86BC25` | Deloitte green — logo dot (sidebar only) |

---

### 9.10 Update ServiceNow Column Names

**Search for:** `SN_EXPECTED_COLUMNS`

```javascript
const SN_EXPECTED_COLUMNS = {
  snEmail:       "Email",
  snRequestNo:   "Request Number",
  snLetterType:  "Letter Type",
  snAddressee:   "Addressed To",
  snCountry:     "Country",
  snPurpose:     "Purpose",
  snTravelFrom:  "Travel Date From",
  snTravelTo:    "Travel Date To",
  snPassport:    "Passport Number",
  snNationality: "Nationality",
  snNotes:       "Additional Notes",
};
```

Change the right-hand values to match your ServiceNow export column headers exactly. The left-hand keys must not be changed.

---

### 9.11 Add, Edit, or Remove Authorised Signatories

**Search for:** `SECTION 6b: AUTHORISED SIGNATORIES`

```javascript
const SIGNATORIES = [
  { name: "Sarah Al Mansoori",  designation: "HR Director" },
  { name: "Khalid Al Hashimi",  designation: "Senior HR Manager" },
  { name: "Divya Nair",         designation: "HR Manager" },
  { name: "Ahmed Al Farsi",     designation: "HR Business Partner" },
];
```

- **Add:** copy any line, paste at the bottom, update name and designation
- **Edit:** change the `name` or `designation` value
- **Remove:** delete the line

Only the name and designation are printed on the letter — no office or location line.

---

### 9.12 Update the Verification Email Address

**Search for:** `empservices@test.com`

Replace with your firm's actual HR verification email address. It appears in the footer note of every letter.

---

## 10. Code Structure Map

```
hr_admin.html
├── SharePoint URL config  → saved in localStorage
├── LOCATIONS              → office validation list
├── EXPECTED_COLUMNS       → HRIS column name mapping
├── processFile()          → reads and parses HRIS Excel
├── autoDetect()           → fuzzy column header matching
├── loadData()             → validates all employee rows
├── renderTable()          → employee data table with status badges
└── publishData()          → exports hris_data.json + shows upload instructions

hr_user.html
├── <style> block
│   ├── SECTION A: App UI styles (sidebar, upload panels, cards, buttons)
│   ├── SECTION B: Letter paper styles (letterhead, tables, footer)
│   └── Confidential banner, letter-meta, addressee-block styles
│
└── <script> block
    ├── SECTION 1: LOCATIONS           → 7 office letterhead definitions
    ├── SECTION 2: LETTER_TYPES        → 11 letter type definitions
    ├── SECTION 3: SAMPLE_EMPLOYEES    → 25 built-in fallback employees
    ├── SECTION 4: EXPECTED_COLUMNS    → HRIS Excel column name mapping
    │             SN_EXPECTED_COLUMNS  → ServiceNow column name mapping
    ├── SECTION 5: TEMPLATES           → All 11 letter body templates
    ├── SECTION 6: ADDRESSEE_PRESETS   → Quick-pick addressee buttons
    ├── SECTION 6b: SIGNATORIES        → Authorised signatory list
    ├── SharePoint integration
    │   ├── loadHrisFromSharePoint()   → fetches hris_data.json by URL
    │   ├── saveSpUrl()                → saves URL to localStorage
    │   └── toggleSpConfig()           → show/hide URL input panel
    ├── SECTION 7: UPLOAD & PARSING
    │   ├── processSourceFile()        → reads ServiceNow Excel
    │   ├── autoDetectMappingFor()     → fuzzy column matching
    │   ├── parseHrisRows()            → converts HRIS rows to objects
    │   ├── parseSnowRows()            → converts SN rows to objects
    │   └── tryMergeAndRender()        → email join + updates list
    ├── SECTION 8: UI RENDERING
    │   ├── renderEmployeeList()       → SN request cards or HRIS cards
    │   ├── collapseUploads()          → hides upload panel after load
    │   └── renderSignatoryDropdown()  → populates signatory select
    ├── SECTION 9: LETTER GENERATION
    │   ├── generateLetter()           → builds full letter HTML
    │   │   ├── Confidential banner
    │   │   ├── Letterhead with logo (192×37px) + office info
    │   │   ├── Request Number (RITM) + date
    │   │   ├── Addressee block
    │   │   ├── Letter body (template)
    │   │   ├── Closing statement (standard across all templates)
    │   │   └── Footer: signatory + verification note
    │   └── selectEmployee()           → auto-fills form from SN data
    ├── SECTION 10: PRINT              → Print / Save PDF window
    └── SECTION 11: EDIT & WORD/PDF
        └── editLetter()               → editable window with Word + PDF download
```

---

## 11. Troubleshooting

**HRIS status shows ❌ red error in the user app**
→ The app could not fetch `hris_data.json` from the configured URL. Check the URL is correct and the file has been uploaded to that location. A fallback HRIS upload zone appears automatically — you can upload the HRIS Excel directly as a workaround.

**HRIS status shows ⚠️ yellow — using sample data**
→ No SharePoint URL has been configured yet. Click "change" next to the URL field, enter the SharePoint folder URL, and click Save.

**A request card shows an email address instead of the employee's name**
→ That ServiceNow row's email does not match any record in the HRIS file. Check for typos or domain mismatches. The card is flagged with a red "No HRIS match" tag.

**The letter type was not auto-selected after clicking a card**
→ The "Letter Type" value in ServiceNow does not closely match any of the 11 built-in template names. Either update the ServiceNow value or manually click the correct letter type button.

**Travel dates are blank in the NOC Tourist / NOC Business letter**
→ The `Travel Date From` and `Travel Date To` columns were empty in the ServiceNow report for that request. Fill them in ServiceNow and re-export, or edit the letter directly using the Edit & Download Word button.

**The column mapping panel appears after upload**
→ Column headers in your Excel file differ from expected names. Use the dropdowns to match them. This only needs to be done once per file format — it is not saved between sessions.

**Employee location badge appears in red**
→ The employee's Location Code does not match any office in the app. Check the spelling (case-sensitive) or add the location — see Section 9.2.

**Salary breakdown shows wrong figures**
→ Ensure the HRIS file has separate columns for Basic Salary, Housing Allowance, and Transport Allowance. If missing, the app uses a 65/25/10 fallback split — see Section 9.5 to adjust.

**Print / PDF shows black boxes instead of coloured table rows**
→ In the print dialog, enable **Background graphics** (Chrome) or **Print background colors and images** (Edge).

**The logo does not appear on the letter**
→ The Deloitte logo is loaded from Wikimedia Commons and requires an internet connection. If offline, the logo area will be blank but all letter content remains unaffected.

**The app doesn't open / shows a blank page**
→ Open in Chrome or Edge. Internet Explorer is not supported.

---

## 12. Frequently Asked Questions

**Do users need to upload the HRIS file every day?**
No. The admin publishes it periodically to SharePoint. The user app fetches it automatically on every open. Users only upload the ServiceNow daily report.

**What if the SharePoint URL fetch fails?**
A red error appears and a fallback upload zone is shown automatically. The user can upload the HRIS Excel directly as a one-off workaround, then contact the admin to resolve the URL issue.

**Can I use the user app without a SharePoint URL?**
Yes — it falls back to 25 built-in sample employees. This is useful for testing but not for live use.

**Does auto-fill from ServiceNow override manual edits?**
Auto-fill only runs once, when a request card is clicked. Any changes made after that are kept until a different card is clicked.

**What if an employee has two open ServiceNow requests?**
Both appear as separate cards in the list (each with their own RITM number). They are processed independently.

**Can multiple HR staff use this simultaneously?**
Yes — each person has their own copy of the HTML file. There is no shared state between users.

**Is employee data stored anywhere?**
No. All data stays in the browser's memory and is cleared when the tab is closed. The SharePoint URL is the only thing saved (in localStorage). No employee data is sent to any server.

**Can I add more than 11 letter types?**
Yes — see Section 9.4.

**Can I add more than 7 offices?**
Yes — see Section 9.2. There is no limit.

**Does the app work on a Mac?**
Yes — open the HTML file in Chrome or Safari on Mac.

**How do I change the verification email address in the footer?**
Search for `empservices@test.com` in the code and replace it — see Section 9.12.

**Who do I contact to make code changes?**
The HTML files are fully self-contained and documented. For significant feature additions, share the file with Claude (claude.ai) for assistance.

---

*This document covers version 6.0 of the HR Letter Generator — Deloitte HR Employee Services.*
