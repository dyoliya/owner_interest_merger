# Owner Interest Merger Tool


The **Owner Interest Merger Tool** is a Python-based preprocessing utility designed to merge **Mineral Owner (MO)** records that belong to the same owner, have the **same number of interests**, but differ in **offer values**.

The tool is intended to be used **outside of the core loading process** to avoid incorrect system behavior where higher offers are flagged as *not latest* and dropped in favor of lower offers.

By standardizing owner and address data, validating critical fields, and applying business-rule-driven merge logic (including special handling for *COMBINED INDIVIDUALS*), the tool safely consolidates eligible MO records while preserving auditability.

The output supports review and validation prior to BUDB loading, helping ensure that the most accurate and complete MO offer data is retained.

---

![Version](https://img.shields.io/badge/version-1.0.0-ffab4c?style=for-the-badge&logo=python&logoColor=white)
![Python](https://img.shields.io/badge/python-3.11%2B-273946?style=for-the-badge&logo=python&logoColor=ffab4c)
![Status](https://img.shields.io/badge/status-active-273946?style=for-the-badge&logo=github&logoColor=ffab4c)

---

## 🚧 Problem Statement / Motivation

This tool was specifically created to address a recurring issue in the handling of **Mineral Owner (MO)** records during the data loading process.

In the existing loading logic:

* Multiple MO records belonging to the **same owner**
* Having the **same # of Interests**
* But associated with **different offer values**

are evaluated in a way that causes **only the lowest offer** to be retained as the *latest* record. Records with higher offers are incorrectly tagged as **“not latest”**, even though they represent valid and current offers.

This behavior results in:

* Loss of higher-value offer information
* Incorrect representation of the most recent or relevant MO data
* Potential downstream impact on valuation, reporting, and decision-making

To prevent this, affected MO records must be **merged outside of the loading process** before ingestion. This tool enables controlled pre-merge processing, ensuring that valid offers are consolidated correctly rather than discarded by default system logic.

The tool assumes that:

* The **Contact Type** is correctly assigned (especially for *COMBINED INDIVIDUALS*, which follow different merge rules)
* The **Well Matching (WM) file** for the corresponding county aligns with the input file, ensuring that the **# of Interests** and **Total Value - Low ($)** used for merging are accurate

Only when these conditions are met should merging be performed.

---

## ✨ Features

### Mineral Owner–Specific Merging

* Designed specifically for **Mineral Owner (MO)** records
* Merges records from the **same owner** with the **same # of Interests** but **different offers**
* Prevents higher offers from being incorrectly dropped during loading

### Data Standardization

* Normalizes **Owner**, **Address**, **City**, **State**, **County**, and **Target State** values
* Expands common abbreviations using predefined dictionaries
* Removes punctuation, spaces, and special characters for reliable matching

### Validation Checks

* Ensures required columns are present
* Flags empty or missing critical fields before processing
* Assumes alignment with the county **Well Matching (WM) file** for interest accuracy

### Intelligent Duplicate Detection

* Uses a composite normalized key consisting of:

  * Owner (Standardized)
  * Address
  * City
  * State
  * County
  * Target State
  * \# of Interests

### Controlled Merge Logic

* **Never merges** records when # of Interests differs
* Applies special logic for **COMBINED INDIVIDUALS** based on prior business rules
* Skips merging when Contact Type conditions are not satisfied

### Financial Field Aggregation

* Safely aggregates offer-related fields:

  * PDP Value ($)
  * Total Value – Low ($)
  * Total Value – High ($)

### Audit & Review Support

* Generates an **Owner Merge Info** field detailing all merged MO records
* Adds a **Merged (Y/N)** indicator
* Adds review remarks when Contact Type adjustments may be required

---
## 🧠 Logic Flow
1. User provides one or more Excel owner files in the input folder.
2. App reads each file and prepares the data for processing.
3. App validates required fields before continuing:
   - Owner (Standardized)
   - \# of Interests
   - Total Value - Low ($)
   - County
   - Target State
  If any required value is missing → file is flagged and processing stops.
4. App standardizes owner and address data to ensure accurate comparisons:
   - Converts text to uppercase
   - Expands common abbreviations (e.g., LLC → LIMITED LIABILITY COMPANY, ST → STREET)
   - Removes punctuation and spaces
5. App groups records that appear to represent the same owner using an exact match of:
   - Standardized Owner
   - Standardized Address
   - Standardized City
   - Standardized State
   - Standardized County
   - Standardized Target State
   - \# of Interests
6. App checks for conflicting interest counts before merging:
   - If the same owner/location has multiple different values of # of Interests
     - the entire group is excluded from merging and kept as-is (Merged = N).
7. App determines whether a group is eligible to merge:
   - If the group contains only one record → no merge (Merged = N)
   - If the group contains two or more records → continue to Contact Type rules
8. If the group contains no COMBINED INDIVIDUALS:
   - All records in the group are merged into one
   - The following values are summed:
     - \# of Interests
     - PDP Value ($)
     - Total Value - Low ($)
     - Total Value - High ($)
9. If the group contains any COMBINED INDIVIDUALS:
   - If the group has fewer than 4 records → no merge
   - If the group has 4 or more records:
     - Only rows that repeat under the same identity are merged
     - Identity is determined by:
       - Standardized Owner
       - First Name
       - Standardized Address, City, State
       - \# of Interests
       - County and Target State
     - Non-repeating rows remain unmerged
10. For every merged record, app creates an Owner Merge Info field that lists:
    - Owner name
    - Owner ID
    - \# of Interests
    - Total Value - Low ($)
    This provides a clear audit trail of what was combined.
11. App adds a reviewer note (Remarks) when all values across the group are identical and none are COMBINED INDIVIDUALS:
    - “Consider modifying CTT to COMBINED INDIVIDUALS”
12. App merges eligible records and produces consolidated rows (Merged = Y).
13. App generates audit outputs:
    - Merged file for downstream loading
    - Duplicate reference file containing only groups that were merged
14. App outputs final files for human review prior to BUDB loading to ensure accuracy before any downstream use.

---

## 📝 Requirements

* Python 3.9 or later
* pandas
* openpyxl (required for Excel I/O)
* re (standard library)
* os (standard library)
* datetime (standard library)

### Input Data Requirements

* Excel files (.xlsx or .xls)

* Required columns:
  * Owner (Standardized)
  * \# of Interests
  * Total Value - Low ($)
  * County
  * Target State

* Optional but supported columns:
  * Owner ID
  * PDP Value ($)
  * Total Value - High ($)
  * Contact Type
  * First Name

---

## 🚀 Installation and Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/owner_interest_merger.git
   cd owner_interest_merger

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt

3. **Folder Structure**

  <pre>project/
  │
  ├─ input/        # Place Excel files here
  ├─ output/       # Generated files will be saved here
  └─ owner_interest_merger.py
  </pre>

4. **Compile the tool**
   ```bash
   pyinstaller --onefile owner_interest_merger.py

---

## 🖥️ User Guide

### Running the Tool

1. Place one or more Excel files into the **input** folder
2. Run the script:
    ```bash
    python owner_interest_merger.py
    ```
3. The tool will automatically:

   * Scan the input folder
   * Process each Excel file
   * Validate required fields
   * Detect and merge eligible duplicates
   * Save results to the output folder


### Output Files

For each input file, two outputs are generated:

#### 1. Duplicate Rows File

* Filename format:

  ```
  duplicate_rows_YYYYMMDD_HHMMSS_<original_filename>.xlsx
  ```
* Contains only rows that were part of a successful merge
* Intended for validation and reviewer reference

#### 2. Merged Output File

* Filename format:

  ```
  merged_output_YYYYMMDD_HHMMSS_<original_filename>.xlsx
  ```
* Contains:

  * Final merged dataset
  * "Merged" flag (Y/N)
    * Indicates whether the row represents a **merged Mineral Owner (MO) record**.
      * `Y` – The row was created by merging multiple MO records
      * `N` – The row was not merged and remains as originally provided
    * Records belonging to groups with **varying `# of Interests`** are always marked `N`
    * Records skipped due to Contact Type rules are marked `N`
  * "Owner Merge Info" audit column
    * Provides a transparent audit trail showing **all MO records combined** into the merged row.
    * Preserves visibility of multiple offers
    * Shows contributing Owner IDs and interest counts
    * Ensures higher offers are not lost during consolidation
    * Format:
      ```
      Owner Name [Owner ID – # of Interests – Total Value Low]
      ```
  * Optional review remarks
    * Provides reviewer guidance when records appear eligible for a **Contact Type (CTT)** adjustment.
    * Possible Value:
      ```
      Consider modifying CTT to COMBINED INDIVIDUALS
      ```
        * This means:
          * Records in the group are identical across key fields
          * Contact Type is not tagged as `COMBINED INDIVIDUALS`
    * **Important:**
      * This is a recommendation only
      * It does not affect merge eligibility or results

### Important Business Rules to Note

* Records with differing **# of Interests** are **never merged**, even if all other fields match
* **COMBINED INDIVIDUALS** records:
  * Require at least 4 rows to be eligible for merging
  * Are merged only when identical rows repeat

⚠️ This tool assists with duplicate detection and consolidation but **does not replace manual review**.

Before output acceptance and BUDB loading:
* Verify merged interest counts
* Cross-check values with corresponding well-matching (WM) files
* Confirm Contact Type accuracy

Incorrect merges can negatively impact downstream reporting, marketing, and sales operations.

---

## 👩‍💻 Credits
- **2025-12-09**: Project created by **Julia** ([@dyoliya](https://github.com/dyoliya))  
- 2025–present: Maintained by **Julia** for **Community Minerals II, LLC**

