---
{"dg-publish":true,"permalink":"/marketing-template-manual/"}
---

## Updates (Change Log)

The following updates or additions have been made to the marketing template. This serves as a running change log:

- Added Power Query to pull in **2 sets of data**, **previous year** and **current year**.
    - (Personal preference: you could put all into a single directory or a single file)
- **New tab** added called **Combined Data**, where all transformations and combinations occur.
    - Combined Data is the export of the Power Query
- Updated all fields using Excel formulas to reference the **Combined Data** tab.
    - This is dynamic and will update as new data is added, removed, or modified
- **Added a Sheet2 table**  for the current year.
    - The table has a formula to remove 1 year from that year to keep incorrect data from populating.

---

## Technical Documentation

There are **three** main Power Queries (PQ) used to automate the process thus far:

1. **File Years**
2. **Combined**
3. **Dummy Data**

Below are the queries, with explanation and additional relevant process details.

---

### File Years (PQ)

**Purpose / Notes**:

- Linked to the **Sheet2** sheet (cells `O10` and `O11`).
    - Cell `O10` contains the named formula `thisyear` (updated annually).
    - Cell `O11` subtracts one year from `O10`.
- This query filters out any potential data not belonging to these two years.
- **No updates** should be needed in the PQ itself.
    - All changes happen in the worksheet cells and during the annual updates.

---

### Combined (PQ)

**Purpose / Notes**:

- Merges data from multiple monthly Prompt files in the specified folder.
- **Extracts** the date from the file name (first 8 characters are assumed to be `MMddyyyy`).
- Combines it with the **Dummy Data** to ensure 24 monthly columns exist.
- Pivots the data so each month becomes its own column in the final table.

---
### Dummy Data (PQ)

**Purpose / Notes**:

- Creates 24 months (12 months each for the **previous** and **current** year).
- Ensures placeholders exist for any missing months in the main data.
- Prevents `#REF!` errors in Excel formulas.

---

## File-Specific Information

Below are the **expected headers** (standardized) across files:

1. **Referring Physician**
2. **Referring Physician Group**
3. **Visit Count for Case**
4. **Referral Received Date**

Each source file needs to match these (or be transformable via PQ).

### webPT File Details

**Purpose / Notes**:

- **Data File Frequency**: Monthly or Yearly
- **Sheet Name**: `Sheet1`
- **File Name Example**: `MMM YYYY Referrals - Web PT Report` (e.g. `Feb 2025 Referrals - Web PT Report`)


#### Frequent Possible Errors File Details

- The file name is **less critical** to the data shape because the WebPT file typically contains date information in the data itself.
- If the sheet name changes (currently `Sheet1`), you will get an error during refresh:
![Pasted image 20250316113319.png](/img/user/Pasted%20image%2020250316113319.png)
- **Header changes** can also cause unexpected issues. For example:
    - If you remove or rename **Referring Physician**, your pivot may produce zero results.
![Pasted image 20250316113534.png](/img/user/Pasted%20image%2020250316113534.png)
    - If the numeric column for **Visit Count for Case** is renamed or changed, the data might fail to load as intended.

: ![Pasted image 20250316113935.png](/img/user/Pasted%20image%2020250316113935.png)

### Prompt File

**Purpose / Notes**:

- **Data File Frequency**: **Monthly**
- **File Name Example**: `mmddyyyy_Referrals by Month Report` (e.g. `02012025_Referrals by Month Report`)
- **Sheet Name**: `Referring Providers

#### Frequent Possible Errors File Details
- If the file name is missing the date or is not in the `MMddyyyy` format, the refresh will fail.
    - ![Pasted image 20250316112842.png](/img/user/Pasted%20image%2020250316112842.png)
- Dialog Box if the file name does not follow expected pattern:
    - ![Pasted image 20250316112906.png](/img/user/Pasted%20image%2020250316112906.png)
- If the date is omitted entirely:
    - ![Pasted image 20250316112932.png](/img/user/Pasted%20image%2020250316112932.png)




### Clinicient

**Purpose / Notes**:

- **Data File Frequency**: **Monthly**
- **File Name Example**: `mmddyyyy__ProCare Referrals - Hunterdon Region January 2025` (e.g. 01012025_ProCare Referrals - Hunterdon Region January 2025
- **Sheet Name**: `Sheet1`
- If the file name is missing the date or is not in the `MMddyyyy` format, the refresh will fail.

#### Frequent Possible Errors File Details
- If the file name is missing the date or is not in the `MMddyyyy` format, the refresh will fail.
    - ![Pasted image 20250316112842.png](/img/user/Pasted%20image%2020250316112842.png)
- Dialog Box if the file name does not follow expected pattern:
    - ![Pasted image 20250316112906.png](/img/user/Pasted%20image%2020250316112906.png)
- If the date is omitted entirely:
    - ![Pasted image 20250316112932.png](/img/user/Pasted%20image%2020250316112932.png)


---

## Update Template Process

1. **Open the Template**. If asked about a security warning, click **Enable Content**.
    - Excel is indicating the file has data connections (this is normal).
![Pasted image 20250316174115.png](/img/user/Pasted%20image%2020250316174115.png)
2. **Update the year** in **Name Manager** (important!).
    - The queries depend on the named formula `thisyear` (cell `O10` in the Sheet2 sheet).
3. Click the **Data** tab in the Ribbon.
![Pasted image 20250316174220.png](/img/user/Pasted%20image%2020250316174220.png)
4. Click **Queries & Connections**, which opens a panel on the right.
![Pasted image 20250316174258.png](/img/user/Pasted%20image%2020250316174258.png)
5. **Double-click** the **Combined Data** query name to open Power Query.
![Pasted image 20250316174321.png](/img/user/Pasted%20image%2020250316174321.png)
6. In the **Applied Steps** panel (on the right), **locate the step** labeled `Source`.
    - Click the **gear icon** next to it to edit the folder path.
![Pasted image 20250316174531.png](/img/user/Pasted%20image%2020250316174531.png)
7. In the dialog box, click **Browse** and navigate to where the new data files are located. Then click **Open**.
    - Confirm the **Folder Path** is correct.
![Pasted image 20250316174554.png](/img/user/Pasted%20image%2020250316174554.png)
8. Click **OK** to apply the change.
9. Close and **Load** the query back to Excel.
![Pasted image 20250316174618.png](/img/user/Pasted%20image%2020250316174618.png)
10. As a precaution, click **Refresh All** and confirm the data loads successfully.
![Pasted image 20250316174640.png](/img/user/Pasted%20image%2020250316174640.png)
    - Monitoring the refresh can be done through **Queries & Connections** panel or the bottom status bar in Excel.
![Pasted image 20250316174702.png](/img/user/Pasted%20image%2020250316174702.png)
---

## Adding New Data

1. **Save the new files** following the naming conventions described above.
    - e.g., `mmddyyyy_Referrals by Month Report` for Prompt, `MMM YYYY Referrals - Web PT Report` for WebPT.
2. **Open the Template**.
3. The template should automatically pull the latest files upon refresh. If not, do the following:
    - Click the **Data** tab in the Ribbon.
![Pasted image 20250316174220.png](/img/user/Pasted%20image%2020250316174220.png)
    - Click **Refresh All**.
![Pasted image 20250316174640.png](/img/user/Pasted%20image%2020250316174640.png)
    - Wait for the refresh to complete (monitor via **Queries & Connections** or the bottom status bar).

---
