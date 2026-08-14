# CareLink Data Formatting

A Streamlit app for turning raw CareLink exports into clean, formatted
output. Three tools live in one app:

- **CareLink Express Data** — upload a raw "Rendered Medicines" report in
  any column layout; the app auto-detects and cleans the data, then
  generates one formatted Excel file **per Patient Source**.
- **CareLink Coordinator Data** — upload a Rendered Medicines file together
  with a Registered Patients file; the app matches them by Patient PIN and
  produces a tracking workbook with a "who got medicine" sheet and a
  "registered but never rendered" sheet.
- **CareLink Checklist Data** — upload the same two files as Coordinator;
  the app pushes a hand-markable follow-up checklist straight into a
  Google Sheet, **one worksheet tab per Patient Source**, covering every
  patient (already served or still awaiting follow-up). Requires a
  one-time Google Sheets setup — see below.

---

## Requirements

- Python 3.9+
- Packages: `streamlit`, `pandas`, `openpyxl`, `gspread`, `google-auth`

```bash
pip install -r requirements.txt
```

## Running it

```bash
streamlit run app.py
```

Then open the URL Streamlit prints (usually `http://localhost:8501`).

### Logo (optional)

Drop a file named exactly `FTCC Head.png` in the same folder as `app.py`
and it'll appear in the header. If it's missing or misnamed, the app falls
back to a placeholder icon and shows a small "Logo not loading" expander
with details on what it looked for and where.

---

## Project structure

```
app.py                 Entry point — page setup, header, tabs. Run this file.
feature_formatter.py   CareLink Express: auto-clean + per-source report generation
feature_merger.py      CareLink Coordinator: two-file PIN-matching + merge
feature_checklist.py   CareLink Checklist: two-file PIN-matching + push to Google Sheets
google_sheets.py        Google Sheets auth + the actual write-to-sheet logic
styles.py                All CSS, plus the small JS that styles the file-upload dropzone
utils.py                 Shared helpers: fuzzy column matching, data cleaning, filename sanitizing
requirements.txt         Python dependencies (used by Streamlit Community Cloud too)
```

Each feature module exposes one `render_*_tab()` function that `app.py`
calls inside its own `st.tabs()` block — this keeps `app.py` itself short
and each feature self-contained.

---

## CareLink Express — how the auto-cleaning works

You don't need to pre-format your file. Upload it as-is; the app will:

1. **Detect each needed field** by column name, tolerant of case, extra
   whitespace, and common alternate names (e.g. `"Cellphone Number"` is
   recognized as Contact Number, `"patient pin"` matches Patient PIN).
   Works whether names are combined (`Patient Name`) or split
   (`Last/First/Middle Name`) — falls back to splitting a combined name if
   separate columns aren't present.
2. **Clean every value**: trims/collapses stray whitespace, strips the
   `.0` pandas adds to numeric-looking PINs, parses any date format into
   `MM/DD/YYYY`, normalizes phone numbers to `09XXXXXXXXX`, uppercases
   ICD10 codes, coerces cost/price/qty fields to real numbers.
3. **Auto-computes Total Cost / Total Price** when missing but Qty
   Dispensed + Cost/Price are present.
4. **Always outputs exactly this column set**, in this order, regardless
   of what the source file looked like:

   ```
   Patient Name, Last Name, First Name, Middle Name, Patient PIN,
   Patient Source, Consultation Date, Rendered Date, End Visit By,
   ICD10 Code, ICD10 Description, Medicine, Medicine Category,
   Qty Prescribed, Qty Dispensed, Cost, Price, Total Cost, Total Price,
   Contact Number, Address, Notes
   ```

5. **Splits the output by Patient Source** — every unique source in the
   upload gets its own workbook, sorted alphabetically (rows with a blank
   source are grouped under "Unspecified Source"). Each workbook has its
   own merged patient blocks and its own totals row.
6. **Filenames are automatic**: `CareLinkExpress-Data-{Source}.xlsx`, with
   unsafe filename characters stripped. There's no manual filename field.

An expander above the results ("Detected column mapping") shows exactly
which raw column was used for each field, so you can sanity-check the
auto-detection before trusting the output.

### Downloading results

- Each source has its own row with a **Download** button.
- **Download all (.zip)** bundles every generated file into one zip.
- With more than 5 sources, a **search box** appears to filter the list by
  source name.
- Check the boxes next to specific sources to reveal a **Download N
  selected (.zip)** button for a partial bundle.

### Removing / replacing the uploaded file

Click the "x" on the uploaded file chip (or drag a different file onto the
uploader) and any previously generated output is cleared immediately — you
won't ever see stale results sitting next to a different (or no) file.
This is tracked via a file "signature" (name + size) stored in
`st.session_state`; whenever it changes, `t1_reports` is reset until you
click **Generate** again.

---

## CareLink Coordinator — how the merge works

1. Upload the **Rendered Medicines** file (its own Cellphone Number / Full
   Address / split-name columns are read directly — no need for a separate
   lookup for patients who did receive medicine).
2. Upload the **Registered Patients** file — used only to find patients
   whose PIN never appears in the Rendered Medicines file.
3. Click **Run Merge**. The output workbook has two sheets:
   - **Rendered Medicines** — one row per dispensed medicine, with
     identity/contact columns merged across each patient's rows.
   - **Registered - Not Rendered** — registered patients with no
     corresponding rendered-medicine record, for follow-up.
4. Both sheets are sorted alphabetically by Last Name, then First Name.
5. The output filename is still manually editable here (unlike Express,
   which is auto-named per source).

Removing/replacing either uploaded file clears any previously generated
output the same way Express does.

---

## CareLink Checklist — how it works, and how to set it up

Unlike the other two tabs, this one doesn't produce a downloadable file —
it writes directly into a Google Sheet.

### What it does

1. Upload the same **Rendered Medicines** + **Registered Patients** files
   as Coordinator.
2. Paste (or use the pre-filled default) the target **Google Sheet URL**.
3. Click **Push to Google Sheets**. For every Patient Source found across
   both files, the app creates or updates a worksheet tab named after that
   source, inside that one spreadsheet — same URL every time, one tab per
   source. Re-running it overwrites those tabs in place rather than piling
   up duplicates.
4. Every patient appears exactly once, in **Last Name, First Name, Middle
   Name** columns, sorted alphabetically by Last Name. Patients who already
   received medicine have their medicines combined into one cell (e.g.
   `Paracetamol, Amoxicillin`); patients still awaiting follow-up show
   blank Medicines/Source. Patients with no detected source land in an
   "Unspecified Source" tab.
5. `1st Contact`, `2nd Contact`, `Prescribed`, and `Packed` are real Google
   Sheets checkboxes; `Consult` is a real dropdown (default options:
   `Pending / Scheduled / Completed / No Show` — see `CONSULT_OPTIONS` in
   `google_sheets.py` to change them).

### One-time setup (required before this tab will work)

The app needs its own Google credentials to write to your sheet — it
can't use your personal Google login. This is a **Service Account**: a
robot identity you create once in Google Cloud, then explicitly grant
Editor access to your spreadsheet.

1. **Create/select a Google Cloud project** at
   https://console.cloud.google.com/
2. **Enable the Sheets API**:
   https://console.cloud.google.com/apis/library/sheets.googleapis.com
   → Enable. (The Drive API is *not* needed — this only writes tabs into
   an existing spreadsheet, never creates new spreadsheet files.)
3. **Create a Service Account**:
   https://console.cloud.google.com/iam-admin/serviceaccounts →
   "+ Create Service Account" → name it anything → Create and Continue →
   skip the optional role steps → Done.
4. **Generate its key**: open the service account → "Keys" tab → "Add
   Key" → "Create new key" → JSON → Create. A `.json` file downloads.
   **Treat this file like a password** — never commit it to a repo or
   share it outside secrets storage.
5. **Share your Google Sheet** with the service account: open the JSON
   file, copy the `client_email` value (looks like
   `name@project.iam.gserviceaccount.com`), then in your actual Google
   Sheet click **Share** → paste that email → role **Editor** → Send.
6. **Add the credentials to Streamlit secrets** — locally, create
   `.streamlit/secrets.toml`; on Streamlit Community Cloud, use the app's
   Settings → Secrets panel. Fill in every field from the JSON file:

   ```toml
   [gcp_service_account]
   type = "service_account"
   project_id = "..."
   private_key_id = "..."
   private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
   client_email = "...@....iam.gserviceaccount.com"
   client_id = "..."
   auth_uri = "https://accounts.google.com/o/oauth2/auth"
   token_uri = "https://oauth2.googleapis.com/token"
   auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
   client_x509_cert_url = "..."

   default_spreadsheet_url = "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit"
   ```

   Keep `private_key` as one line with literal `\n` characters exactly as
   they appear in the JSON — don't reformat it into real line breaks.

Until this is done, the Checklist tab shows a plain "Google Sheets isn't
connected yet" message instead of the upload UI — it won't error out, it
just won't do anything.

### If it fails after setup

The app's error message includes the service account's email as a hint,
since **the #1 real-world failure is forgetting to share the sheet with
it**. Double-check Step 5 above first. A "Technical details" expander on
the error also shows the raw exception if you need to dig further.

---

## Known constraints

- **One header row per file.** Like any real single-sheet Excel export, a
  column's meaning is fixed for the whole file — the auto-detection can't
  read row 1 as one naming convention and row 50 as another within the
  *same* column header.
- If a needed field genuinely isn't present anywhere in the upload, that
  column comes out blank rather than guessed — check the "Detected column
  mapping" expander if a report looks incomplete.

---

## Troubleshooting: "I updated the code but the output didn't change"

This is almost always a **stale running process**, not a code bug.
Streamlit reruns your top-level script on every interaction, but it does
not reliably reload code from *imported* modules (like
`feature_formatter.py`) within an already-running process — Python caches
imported modules in memory for the life of that process.

1. Confirm the file on the actual running server has your latest code:
   ```bash
   grep -c "_normalize_to_output_schema" feature_formatter.py
   ```
   If this prints `0`, that server is not running the file you think it is
   — check for multiple deployments, wrong branch, or a stale build.
2. If it prints `2` or more but the bug persists, the process itself is
   stale. Fully kill it and restart:
   ```bash
   ps aux | grep -i streamlit   # find the PID(s)
   kill -9 <PID>
   lsof -i :8501                # confirm nothing's still listening
   streamlit run app.py
   ```
3. On Streamlit Community Cloud, use the dashboard's **Reboot app**
   button — a git push alone doesn't guarantee the old process was torn
   down first.
4. Test in a fresh/incognito browser tab to rule out any client-side
   caching of a previous response.

---

## Extending it

- **Add a new fuzzy-matched field**: add a candidate list to
  `_detect_columns()` in `feature_formatter.py` (or `_detect_med_columns()`
  in `feature_merger.py`), then reference it in the corresponding
  normalize/build function.
- **Change what counts as "clean"**: the per-type cleaning functions
  (`clean_str`, `clean_pin`, `clean_date`, `clean_number`,
  `split_name_fallback`) all live in `utils.py` and are shared by both
  features.
- **Change the output column set**: edit `OUTPUT_COLUMNS`,  `MERGE_COLS`,
  and `CENTERED_COLS` at the top of `feature_formatter.py`.