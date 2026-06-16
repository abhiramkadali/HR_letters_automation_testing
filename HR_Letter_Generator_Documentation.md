# HR Letter Generator — Full Documentation

**Tool:** `hr_letter_generator.html`
**Version:** 5.0
**Prepared by:** HR Employee Services
**Last Updated:** June 2026 (v5.0 — final build)

---

## Table of Contents

1. [What This Tool Does](#1-what-this-tool-does)
2. [How the Two-Source System Works](#2-how-the-two-source-system-works)
3. [How to Open and Use the App](#3-how-to-open-and-use-the-app)
4. [How to Upload Data](#4-how-to-upload-data)
   - 4.1 [Uploading the HRIS File](#41-uploading-the-hris-file)
   - 4.2 [Uploading the ServiceNow File](#42-uploading-the-servicenow-file)
   - 4.3 [What Happens After Both Files Are Loaded](#43-what-happens-after-both-files-are-loaded)
5. [Excel File Formats — Required Columns](#5-excel-file-formats--required-columns)
   - 5.1 [HRIS File — Employee Master Data](#51-hris-file--employee-master-data)
   - 5.2 [ServiceNow File — Letter Requests](#52-servicenow-file--letter-requests)
   - 5.3 [Valid Location Codes](#53-valid-location-codes)
   - 5.4 [Sample Files](#54-sample-files)
6. [How to Save a Letter as PDF](#6-how-to-save-a-letter-as-pdf)
7. [How to Make Changes to the Code](#7-how-to-make-changes-to-the-code)
   - 7.1 [Update an Office Address or Phone Number](#71-update-an-office-address-or-phone-number)
   - 7.2 [Add a New Office Location](#72-add-a-new-office-location)
   - 7.3 [Edit Wording in a Letter Template](#73-edit-wording-in-a-letter-template)
   - 7.4 [Add a New Letter Type](#74-add-a-new-letter-type)
   - 7.5 [Change the Salary Split Percentages](#75-change-the-salary-split-percentages)
   - 7.6 [Add or Edit Addressee Quick-Pick Options](#76-add-or-edit-addressee-quick-pick-options)
   - 7.7 [Change the Company Name or Logo](#77-change-the-company-name-or-logo)
   - 7.8 [Change the Signatory Title](#78-change-the-signatory-title)
   - 7.9 [Change Brand Colours](#79-change-brand-colours)
   - 7.10 [Update ServiceNow Column Names](#710-update-servicenow-column-names)
   - 7.11 [Add, Edit, or Remove Authorised Signatories](#711-add-edit-or-remove-authorised-signatories)
8. [Code Structure Map](#8-code-structure-map)
9. [Troubleshooting](#9-troubleshooting)
10. [Frequently Asked Questions](#10-frequently-asked-questions)

---

## 1. What This Tool Does

The HR Letter Generator is a self-contained web application that runs entirely inside a single HTML file. No internet connection, no login, no server, and no installation is required.

From version 5.0 onwards, the app operates on **two data sources** uploaded by the HR user:

| Source | File | Purpose |
|---|---|---|
| **HRIS** | `HRIS_Sample.xlsx` (or your own export) | Employee master data — salary, designation, bank details, location |
| **ServiceNow** | `ServiceNow_Sample.xlsx` (or your own report) | Today's active letter requests — who needs what letter and why |

The two files are **joined automatically on email address.** The employee list in the sidebar shows **only the employees who have submitted a letter request via ServiceNow**, not the entire company roster. This means HR can open the tool, upload two files, and immediately see exactly which requests to process — without searching through hundreds of employees.

The app allows HR to:

- Upload HRIS and ServiceNow exports and merge them in one click
- See only active letter requestors in the employee list (not the full company)
- Have the letter type, addressee, and purpose auto-filled from the ServiceNow ticket
- Choose from 9 standard letter types
- Preview the fully formatted letter instantly with the correct office letterhead
- Print or save as PDF

The correct office letterhead is applied automatically based on the employee's work location. With 7 offices and 9 letter types, the app handles 63 combinations without any manual formatting.

---

## 2. How the Two-Source System Works

```
HRIS Export                    ServiceNow Report
(all employees)                (today's requests only)
      │                               │
      │   JOIN on Email Address       │
      └──────────────┬────────────────┘
                     ▼
         Merged employee list
         (only requestors shown)
                     │
          ┌──────────┴──────────┐
          │  Auto-filled form   │
          │  Letter type        │
          │  Addressee          │
          │  Purpose            │
          └──────────┬──────────┘
                     ▼
            Generate → PDF
```

**The join key is the employee's work email address.** Every row in the ServiceNow report must have an email that matches a row in the HRIS file. If a ServiceNow row's email has no HRIS match, it still appears in the list but is flagged in red so HR can investigate.

**ServiceNow fields take priority over HRIS for overlapping data.** Specifically, if the ServiceNow report contains a Passport Number or Nationality (e.g. because the employee submitted an updated passport with their request), those values override what is in HRIS for that letter only.

---

## 3. How to Open and Use the App

**Step 1** — Save `hr_letter_generator.html` anywhere on your computer (Desktop, Documents, etc.)

**Step 2** — Double-click the file to open it in your browser (Chrome or Edge recommended).

**Step 3** — Follow the steps in the left panel:

| Step | Action |
|------|--------|
| 1 | Upload your **HRIS file** into the HRIS upload zone |
| 2 | Upload your **ServiceNow report** into the ServiceNow upload zone |
| 3 | The employee list refreshes automatically — only active requestors are shown |
| 4 | Click a request card to select it — letter type, purpose, and addressee are auto-filled |
| 5 | Adjust any auto-filled fields if needed |
| 6 | Set the letter date (defaults to today) |
| 7 | Click **Generate Letter →** |
| 8 | Review the preview, then click **Print / Save PDF** |

> **Note:** You can upload the files in either order. The merge fires as soon as both are present. If you only upload the HRIS file, the full employee list is shown (legacy mode). If you only upload the ServiceNow file, request rows are shown with email addresses as placeholders until the HRIS file is added.

> **Offline use:** The app works fully offline. The only feature that requires internet is loading the display fonts (Inter, Playfair Display) from Google Fonts. All letter content and data logic work without a connection.

---

## 4. How to Upload Data

### 4.1 Uploading the HRIS File

The HRIS block is labelled with a blue **HRIS** badge in the sidebar.

- Click the upload zone or drag and drop your HRIS `.xlsx` export onto it
- The app reads the file instantly and shows a green status message with the employee count
- If column names differ from expected (e.g. your HRIS uses "Staff ID" instead of "Employee ID"), a **mapping panel** appears — use the dropdowns to match your columns, then click **Apply & Load HRIS**

### 4.2 Uploading the ServiceNow File

The ServiceNow block is labelled with an amber **ServiceNow** badge directly below the HRIS block.

- Click the upload zone or drag and drop your ServiceNow `.xlsx` report onto it
- The app reads the file and shows a status message with the request count
- If column names differ, a mapping panel appears for ServiceNow columns separately
- Only rows with a valid email address are loaded (rows without email are skipped)

### 4.3 What Happens After Both Files Are Loaded

Once both files are uploaded, the app:

1. **Joins** every ServiceNow row to its HRIS record via email address
2. **Replaces** the employee list with only the matched requestors
3. Shows a **merge banner** below the upload panels:
   - ✅ Green — all requests matched to HRIS records
   - ⚠️ Amber — some requests could not be matched (those rows appear in red in the list)
4. **Auto-fills** the form when you click a request card:
   - Letter type is selected automatically if the ServiceNow "Letter Type" matches a known template
   - Purpose is filled from the ServiceNow "Purpose" column
   - Addressee is filled from the ServiceNow "Addressed To" column

You can re-upload either file at any time — the list refreshes immediately.

---

## 5. Excel File Formats — Required Columns

### 5.1 HRIS File — Employee Master Data

Column order does not matter — only the header names matter. The app uses fuzzy matching, so minor variations (e.g. "Emp ID" instead of "Employee ID") will be caught automatically or prompted via the mapping panel.

| Column Header | Required | Description | Example |
|---|---|---|---|
| `Employee ID` | Recommended | Unique employee code | EMP001 |
| `Name` | **Yes** | Full name as it appears on the letter | Mohammed Al Rashidi |
| `Email` | **Yes** | Work email — **join key** for matching ServiceNow rows | m.rashidi@company.com |
| `Designation` | Yes | Job title | Senior Financial Analyst |
| `Department` | Yes | Department name | Finance |
| `Location Code` | Yes | Must match an office code (see Section 5.3) | DXB |
| `Join Date` | Yes | Date of joining | 15 March 2019 |
| `Nationality` | Yes | Nationality | Emirati |
| `Passport Number` | Yes | Passport number | A12345678 |
| `Basic Salary` | Yes | Basic salary component | AED 12,025 |
| `Housing Allowance` | Yes | Housing allowance | AED 4,625 |
| `Transport Allowance` | Yes | Transport allowance | AED 1,850 |
| `Total Salary` | Yes | Total monthly salary | AED 18,500 |
| `Bank` | Yes | Employee's bank name | Emirates NBD |
| `IBAN` | Yes | IBAN number | AE07 0331 2345 6789 0123 456 |
| `Address` | Yes | Residential address | Villa 12, Al Barsha 2, Dubai, UAE |

> **Salary tip:** If your HRIS only exports a Total Salary and no breakdown, the app will auto-calculate Basic (65%), Housing (25%), and Transport (10%) as a fallback. See Section 7.5 to change these percentages.

### 5.2 ServiceNow File — Letter Requests

This is the report you export from ServiceNow containing the employees who have submitted letter requests. Only employees present in this file will appear in the sidebar list.

| Column Header | Required | Description | Example |
|---|---|---|---|
| `Email` | **Yes** | Work email — **join key** — must match HRIS | m.rashidi@company.com |
| `Request Number` | Recommended | ServiceNow ticket / RITM number | RITM0012301 |
| `Letter Type` | Recommended | Requested letter type | Salary Certificate |
| `Addressed To` | Optional | Who the letter should go to | The Branch Manager |
| `Country` | Optional | Country of the recipient entity — appended to the addressee line | UAE |
| `Purpose` | Optional | Reason stated in the ticket | bank account upgrade |
| `Travel Date From` | Optional | Travel start date (visa/NOC letters) | 10 July 2026 |
| `Travel Date To` | Optional | Travel end date | 25 July 2026 |
| `Passport Number` | Optional | Overrides HRIS passport if provided | A12345678 |
| `Nationality` | Optional | Overrides HRIS nationality if provided | Emirati |
| `Additional Notes` | Optional | Free-text notes from the requestor | Please issue urgently |

> **Letter Type matching:** The value in the "Letter Type" column is matched against the app's 9 built-in template names using fuzzy matching. For reliable auto-selection, use names close to: `Salary Certificate`, `Salary Transfer Letter`, `Employment Certificate`, `No Objection Certificate`, `Experience Letter`, `Visa Support Letter`, `Bank Account Opening`, `Address Proof Letter`, `Reference Letter`.

### 5.3 Valid Location Codes

The `Location Code` in the HRIS file must exactly match one of the codes below (case-sensitive). If an employee has an unrecognised code, their card appears with a red location badge.

| Code | Office |
|------|--------|
| DXB | Dubai – Head Office |
| AUH | Abu Dhabi |
| SHJ | Sharjah |
| AJM | Ajman |
| RAK | Ras Al Khaimah |
| UAQ | Umm Al Quwain |
| FUJ | Fujairah |

### 5.4 Sample Files

Two ready-to-use sample files are provided alongside this documentation:

| File | Contents |
|---|---|
| `HRIS_Sample.xlsx` | 10 sample employees across all 7 UAE offices with all required columns |
| `ServiceNow_Sample.xlsx` | 10 sample letter requests covering a range of letter types; one row intentionally has no HRIS match to demonstrate the unmatched row behaviour |

Upload both files into the app to see a fully working example before connecting your live data.

---

## 6. How to Save a Letter as PDF

1. After generating a letter, click **Print / Save PDF** in the top toolbar
2. A print preview window opens in a new browser tab
3. In the print dialog:
   - Set **Destination** to **Save as PDF**
   - Set **Paper size** to A4
   - Enable **Background graphics** (so the dark table rows print correctly)
4. Click **Save**

> **Tip:** In Chrome, press `Ctrl + P` after the print window opens. In Edge, the same shortcut works.

---

## 7. How to Make Changes to the Code

All changes are made by opening `hr_letter_generator.html` in **Notepad** (right-click → Open with → Notepad) and editing the relevant section.

Use **Ctrl + F** to search for the text shown in each instruction below.

---

### 7.1 Update an Office Address or Phone Number

**Search for:** `SECTION 1: OFFICE LOCATIONS`

You will see a list like this:

```javascript
{ id:"DXB", name:"Dubai – Head Office", address:"Level 12, Dubai World Trade Centre...", phone:"+971 4 123 4567", email:"hr.dubai@company.com" },
```

Find the office by its `id` and change the relevant field. Save the file.

---

### 7.2 Add a New Office Location

**Search for:** `SECTION 1: OFFICE LOCATIONS`

Copy any existing line and paste it below the last entry. Update all fields:

```javascript
{ id:"XYZ", name:"New Office Name", address:"Full address", phone:"+971 X XXX XXXX", email:"hr.xyz@company.com" },
```

> **Important:** The `id` must exactly match the `Location Code` value in your HRIS file for employees at that office.

---

### 7.3 Edit Wording in a Letter Template

**Search for:** `SECTION 5: LETTER TEMPLATES`

Each template is labelled clearly, e.g.:
```
// ── TEMPLATE 1: Salary Certificate
// ── TEMPLATE 3: Employment Certificate
```

Find the template you want to edit and change the paragraph text. Employee data is inserted using placeholders — do not remove these.

**Available placeholders:**

| Placeholder | What it inserts |
|---|---|
| `${emp.name}` | Employee full name |
| `${emp.designation}` | Job title |
| `${emp.department}` | Department |
| `${emp.joinDate}` | Date of joining |
| `${emp.nationality}` | Nationality (ServiceNow value used if provided) |
| `${emp.passportNo}` | Passport number (ServiceNow value used if provided) |
| `${emp.salary}` | Total monthly salary |
| `${emp.basic}` | Basic salary |
| `${emp.housing}` | Housing allowance |
| `${emp.transport}` | Transport allowance |
| `${emp.bank}` | Bank name |
| `${emp.iban}` | IBAN number |
| `${emp.address}` | Residential address |
| `${loc.name}` | Office name |
| `${loc.email}` | Office HR email |
| `${purpose}` | Purpose (auto-filled from ServiceNow, editable) |
| `${addressee}` | Addressee (auto-filled from ServiceNow, editable) |

---

### 7.4 Add a New Letter Type

This requires two steps:

**Step 1 — Register the letter type:**
Search for `SECTION 2: LETTER TYPES` and add a new entry:

```javascript
{ id:"your_letter_id", name:"Your Letter Name", description:"Short description" },
```

**Step 2 — Add the template body:**
Search for `SECTION 5: LETTER TEMPLATES` and add a new entry before the closing `};`:

```javascript
your_letter_id: (emp, loc, purpose, addressee) => `
  <p class="salutation">To: ${addressee},</p>
  <p class="subject-line">SUBJECT: YOUR SUBJECT HERE</p>
  <p>Body text using ${emp.name}, ${emp.designation}, etc.</p>
`,
```

> The `id` in both steps must match exactly. The `name` you use here is also what ServiceNow's "Letter Type" column is matched against for auto-selection.

---

### 7.5 Change the Salary Split Percentages

**Search for:** `TEMPLATE 1: Salary Certificate`

You will see:
```javascript
const basicAmt = emp.basic || "AED " + Math.round(total * 0.65).toLocaleString();
const housAmt  = emp.housing || "AED " + Math.round(total * 0.25).toLocaleString();
const tranAmt  = emp.transport || "AED " + Math.round(total * 0.10).toLocaleString();
```

Change `0.65`, `0.25`, `0.10` to your desired percentages. They must add up to 1.0 (100%).

> These fallback percentages only apply when the HRIS file has no separate Basic/Housing/Transport columns. If those columns exist, their actual values are always used instead.

---

### 7.6 Add or Edit Addressee Quick-Pick Options

**Search for:** `SECTION 6: ADDRESSEE PRESETS`

```javascript
const ADDRESSEE_PRESETS = [
  { label: "Bank",     value: "The Branch Manager" },
  { label: "Embassy",  value: "The Consulate General / Embassy" },
  ...
];
```

- `label` → short text shown on the button in the sidebar
- `value` → what appears on the printed letter

Add a new line to add an option; change `label` or `value` to edit an existing one.

---

### 7.7 Change the Company Name or Logo

**Search for:** `Deloitte<span style="color:#86BC25;">`

This appears in two places — the sidebar header and the letter header. Update the text in both.

To change the green dot colour, replace `#86BC25` with any hex colour code.

Also search for `Deloitte —` to update the footer signature line.

---

### 7.8 Change the Signatory Title

**Search for:** `Authorised Signatory`

Change the text to the desired name or title. It appears in both the live preview and the print window.

---

### 7.9 Change Brand Colours

**Search and replace in Notepad** (`Ctrl + H`):

| Current value | What it controls |
|---|---|
| `#0f172a` | Dark navy — sidebar background, borders, table totals, headings |
| `#2563eb` | Blue — Generate button, active selections |
| `#86BC25` | Deloitte green — logo dot |

Replace with your preferred hex colour codes.

---

### 7.10 Update ServiceNow Column Names

If your ServiceNow report uses different column header names than the defaults, you can update the expected names in the code so the app auto-detects them without showing the mapping panel every time.

**Search for:** `SN_EXPECTED_COLUMNS`

You will see:

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

Change the right-hand values (in quotes) to exactly match the column headers in your ServiceNow export. The left-hand keys must not be changed.

> The same section for HRIS columns is called `EXPECTED_COLUMNS` and follows the same pattern.

---

### 7.11 Add, Edit, or Remove Authorised Signatories

**Search for:** `SECTION 6b: AUTHORISED SIGNATORIES`

You will see:

```javascript
const SIGNATORIES = [
  { name: "Sarah Al Mansoori",  designation: "HR Director" },
  { name: "Khalid Al Hashimi",  designation: "Senior HR Manager" },
  { name: "Divya Nair",         designation: "HR Manager" },
  { name: "Ahmed Al Farsi",     designation: "HR Business Partner" },
];
```

- **To add a signatory:** copy any existing line, paste it at the bottom of the list, and update the `name` and `designation` fields.
- **To edit a signatory:** change the `name` or `designation` value directly.
- **To remove a signatory:** delete their line from the list.

The selected signatory's name and designation print at the bottom of every letter. The office name and location are not printed beneath them — only the name and designation appear.

---

## 8. Code Structure Map

```
hr_letter_generator.html
│
├── <style> block
│   ├── SECTION A: App UI styles (sidebar, buttons, layout)
│   │   └── Dual-upload panel styles (upload-source-block, source-badge,
│   │       sn-card, merge-banner, source-status)
│   └── SECTION B: Letter paper styles (letterhead, tables, footer)
│
└── <script> block
    ├── SECTION 1: LOCATIONS          → Office letterhead data
    ├── SECTION 2: LETTER_TYPES       → 9 letter type definitions
    ├── SECTION 3: SAMPLE_EMPLOYEES   → Built-in demo data (used before any upload)
    ├── SECTION 4: EXPECTED_COLUMNS   → HRIS Excel column name mapping
    │             SN_EXPECTED_COLUMNS → ServiceNow Excel column name mapping
    ├── SECTION 5: TEMPLATES          → All 9 letter body templates
    ├── SECTION 6: ADDRESSEE_PRESETS  → Quick-pick addressee options
    ├── SECTION 6b: SIGNATORIES       → Authorised signatory names and designations
    ├── SECTION 7: DUAL-SOURCE UPLOAD → File reading, parsing, join logic
    │   ├── processSourceFile()       → Parses either HRIS or SN file
    │   ├── autoDetectMappingFor()    → Fuzzy column header matching
    │   ├── parseHrisRows()           → Converts HRIS rows to employee objects
    │   ├── parseSnowRows()           → Converts SN rows to request objects
    │   └── tryMergeAndRender()       → Joins both sources and updates the list
    ├── SECTION 8: UI RENDERING       → Employee list, buttons, search
    │   └── renderEmployeeList()      → Shows SN request cards or plain HRIS cards
    ├── SECTION 9: LETTER GENERATION  → Builds the full letter HTML
    │   └── selectEmployee()          → Auto-fills form from SN data on selection
    └── SECTION 10: PRINT             → Opens print/PDF window
```

---

## 9. Troubleshooting

**The app opens but shows no employees**
→ Upload your HRIS file using the HRIS upload zone. The built-in sample data is for demo purposes only and shows until the first file is uploaded.

**A request card shows the email address instead of the employee's name**
→ That ServiceNow row's email does not match any record in the HRIS file. Check that the email in the ServiceNow report exactly matches the one in HRIS (including spelling and domain). The card is also shown with a red "No HRIS match" tag.

**The merge banner shows unmatched requests**
→ One or more ServiceNow rows could not be joined to HRIS. Common causes: typo in the email address, employee is a contractor not in HRIS, or the HRIS export is out of date. Fix the email in either file and re-upload.

**The letter type was not auto-selected after clicking a card**
→ The "Letter Type" value in your ServiceNow report does not closely match any of the 9 built-in template names. Either update the ServiceNow value to match (e.g. "Salary Certificate"), or manually click the correct letter type button after selecting the employee.

**The column mapping panel appears after upload**
→ The app could not recognise some column headers automatically. Use the dropdowns to match your columns to the expected fields. This only needs to be done once per file format — the mapping is not saved between sessions.

**Employee location badge appears in red**
→ The employee's Location Code in the HRIS file doesn't match any office in the app. Check the spelling (it is case-sensitive) or add the location in Section 7.2.

**Salary breakdown shows wrong figures**
→ Make sure your HRIS file has separate columns for Basic Salary, Housing Allowance, and Transport Allowance. If these are missing, the app uses a 65/25/10 split as a fallback. See Section 7.5 to adjust those percentages.

**Print / PDF shows black boxes instead of coloured table rows**
→ In the print dialog, enable **Background graphics** (Chrome) or **Print background colors and images** (Edge).

**The app doesn't open / shows a blank page**
→ Try opening in Chrome or Edge. Internet Explorer is not supported.

**Fonts look different when offline**
→ The app loads fonts from Google Fonts. Offline, it falls back to system fonts — the letter content is unaffected.

---

## 10. Frequently Asked Questions

**Can I use the app without a ServiceNow file?**
Yes. If you only upload the HRIS file, the app works exactly as it did in version 4.0 — the full employee list is shown and you select manually. The ServiceNow upload is optional but recommended for request-driven workflows.

**Can I upload the files in any order?**
Yes. Upload HRIS first, then ServiceNow — or the other way around. The merge runs automatically as soon as both files are present.

**Does auto-fill from ServiceNow override my manual edits?**
Auto-fill only runs once, when you click a request card. If you change the letter type, purpose, or addressee after that, your changes are kept. Clicking a different employee card will re-apply auto-fill for that new selection.

**What if an employee has two open ServiceNow requests at the same time?**
Both requests appear as separate cards in the list (each with their own request number), even if they belong to the same employee. You process them independently.

**Can multiple HR staff use this at the same time?**
Yes — each person has their own copy of the HTML file on their device. There is no shared server or shared state.

**Is employee data stored anywhere?**
No. All data stays in the browser's memory and is cleared when the tab is closed. Nothing is sent to any server.

**Can I add more than 9 letter types?**
Yes — see Section 7.4 for step-by-step instructions.

**Can I add more than 7 offices?**
Yes — see Section 7.2. There is no limit.

**Does the app work on a Mac?**
Yes — open the HTML file in Chrome or Safari on Mac.

**What if our salary structure changes?**
Update the HRIS file with new salary figures and re-upload. The app immediately reflects the new data.

**Who do I contact to make code changes?**
The HTML file is fully self-contained and documented. Any changes follow the instructions in Section 7. For significant feature additions, the file can be shared with Claude (claude.ai) for assistance.

---

*This document covers version 5.0 of the HR Letter Generator — Deloitte HR Employee Services.*
