# -------------------------ABOUT --------------------------

# pyinstaller --onefile uw_received_offers_tracker.py
# Tool: Owner Interest Merger Tool
# Developer: dyoliya
# Created: 2025-12-09

# © 2025 dyoliya. All rights reserved.

# ---------------------------------------------------------

import pandas as pd
import re
import os
from datetime import datetime

# ---------------------------------------------------------
# Address Abbreviation Dictionary (you can expand this)
# Example: {"ST": "STREET", "RD": "ROAD"}
# ---------------------------------------------------------
ADDRESS_ABBREVIATIONS = {
    "APT":  "APARTMENT",
    "AVE":  "AVENUE",
    "BLVD":  "BOULEVARD",
    "CR":  "CIRCLE",
    "CT":  "COURT",
    "DR":  "DRIVE",
    "E":  "EAST",
    "HWY":  "HIGHWAY",
    "LN":  "LANE",
    "N":  "NORTH",
    "NE":  "NORTHEAST",
    "NW":  "NORTHWEST",
    "PKWY":  "PARKWAY",
    "PO":  "P.O.",
    "RD":  "ROAD",
    "S":  "SOUTH",
    "SE":  "SOUTHEAST",
    "ST":  "STREET",
    "STE":  "SUITE",
    "SW":  "SOUTHWEST",
    "TRL":  "TRAIL",
    "W":  "WEST",
}

OWNER_ABBREVIATIONS = {
    "CO": "COMPANY",
    "CORP": "CORPORATION",
    "EST": "ESTATE",
    "FAM": "FAMILY",
    "FMLY": "FAMILY",
    "INC": "INCORPORATED",
    "IRR": "IRREVOCABLE",
    "IRREV": "IRREVOCABLE",
    "IRRV": "IRREVOCABLE",
    "IRRVCABLE": "IRREVOCABLE",
    "IRV": "IRREVOCABLE",
    "LIV": "LIVING",
    "LLC": "LIMITED LIABILITY COMPANY",
    "LP": "LIMITED PARTNERSHIP",
    "LTD": "LIMITED",
    "LVG": "LIVING",
    "REV": "REVOCABLE",
    "REVC": "REVOCABLE",
    "REVOC": "REVOCABLE",
    "RLT": "REVOCABLE LIVING TRUST",
    "TR": "TRUST",
    "TRST": "TRUST",
    "TRT": "TRUST",
    "TRTEE": "TRUSTEE",
    "TST": "TRUST",
    "TSTE": "TRUSTEE",
    "TSTEE": "TRUSTEE",
    "TSTEE": "TRUSTEE",
    "TSTEES": "TRUSTEES",
    "TTEE": "TRUSTEE",
}

def normalize_text(value, is_owner=False):
    if pd.isna(value):
        return ""
    value = str(value).upper()

    # If owner field, apply OWNER abbreviations
    if is_owner:
        for k, v in OWNER_ABBREVIATIONS.items():
            pattern = r'\b' + re.escape(k) + r'\b'
            value = re.sub(pattern, v, value)

    # Apply address abbreviations ONLY if not owner
    if not is_owner:
        for k, v in ADDRESS_ABBREVIATIONS.items():
            pattern = r'\b' + re.escape(k) + r'\b'
            value = re.sub(pattern, v, value)

    # Remove special characters and spaces
    value = re.sub(r'[^A-Z0-9]', '', value)

    return value

# ---------------------------------------------------------
# INITIAL VALIDATION
# ---------------------------------------------------------
def validate_required_fields(df):
    required_cols = [
        "Owner (Standardized)",
        "# of Interests",
        "Total Value - Low ($)",
        "County",
        "Target State"
    ]

    errors = []

    for col in required_cols:
        if col not in df.columns:
            errors.append(f"Missing required column: {col}")
            continue

        for idx, val in df[col].items():
            if str(val).strip() == "" or pd.isna(val):
                errors.append(f"❌ Empty value in `{col}` at row {idx + 2}")

    return errors

# ---------------------------------------------------------
# MAIN FUNCTION
# ---------------------------------------------------------
def run_owner_interest_merger(df, file_name, output_folder):

    # ---- Pre-clean: drop fully empty or whitespace-only rows
    df = df.replace(r'^\s*$', pd.NA, regex=True)  # turn whitespace into NA
    df = df.dropna(how='all')  # drop rows that are completely empty
    df = df.reset_index(drop=True)

    original_row_count = len(df)

    # ---- Step 1: Validation
    validation_errors = validate_required_fields(df)

    if validation_errors:
        raise ValueError(
            f"Validation failed for {file_name}:\n" + "\n".join(validation_errors)
        )

    # ---- Step 2: Create normalized keys for matching
    df_work = df.copy()

    df_work["__norm_owner"] = df_work["Owner (Standardized)"].apply(lambda x: normalize_text(x, is_owner=True))
    df_work["__norm_address"] = df_work["Address"].apply(normalize_text)
    df_work["__norm_city"] = df_work["City"].apply(normalize_text)
    df_work["__norm_state"] = df_work["State"].apply(normalize_text)
    df_work["__norm_county"] = df_work["County"].apply(normalize_text)
    df_work["__norm_target_state"] = df_work["Target State"].apply(normalize_text)
    df_work["__norm_interests"] = df_work["# of Interests"].astype(str).apply(normalize_text)

    # Key used for duplicates
    df_work["__dup_key"] = (
        df_work["__norm_owner"] + "|" +
        df_work["__norm_address"] + "|" +
        df_work["__norm_city"] + "|" +
        df_work["__norm_state"] + "|" +
        df_work["__norm_county"] + "|" +
        df_work["__norm_target_state"] + "|" +
        df_work["__norm_interests"]
    )

    # ---- NEW: Filter groups with varying # of Interests
    interest_filter_key = (
        df_work["__norm_owner"] + "|" +
        df_work["__norm_address"] + "|" +
        df_work["__norm_city"] + "|" +
        df_work["__norm_state"] + "|" +
        df_work["__norm_county"] + "|" +
        df_work["__norm_target_state"]
    )

    df_work["__interest_group_key"] = interest_filter_key

    # Identify groups where # of Interests is not all the same
    def has_varying_interests(sub_df):
        return sub_df["# of Interests"].nunique() > 1

    varying_interest_groups = (
        df_work.groupby("__interest_group_key", dropna=False)
        .filter(has_varying_interests)["__interest_group_key"]
        .unique()
    )

    # Mark rows that should be skipped for merging
    df_work["__skip_merge"] = df_work["__interest_group_key"].isin(varying_interest_groups)

    # ---- Step 3: Detect duplicates
    duplicate_mask = df_work.duplicated(subset="__dup_key", keep=False)
    duplicates_df = df.loc[duplicate_mask].copy()

    # ---- Step 4: Merge duplicates
    merge_fields = [
        "# of Interests",
        "PDP Value ($)",
        "Total Value - Low ($)",
        "Total Value - High ($)"
    ]

    def get_owner_with_details(owner, owner_id, num_interests, total_value_low):
        if pd.isna(owner_id) or str(owner_id).strip() == "":
            owner_id_val = "No Owner ID"
        else:
            owner_id_str = str(owner_id).strip()
            owner_id_val = re.sub(r"\.0$", "", owner_id_str)

        interests_val = str(num_interests) if not pd.isna(num_interests) else "0"
        value_low_val = str(total_value_low) if not pd.isna(total_value_low) else "0"
        return f"{owner} [{owner_id_val} - {interests_val} - {value_low_val}]"

    def safe_float(val):
        try:
            return float(str(val).replace(",", "").strip())
        except:
            return 0.0

    merged_rows = []
    merged_flag_col = "Merged"
    merged_groups_keys = set()  # Track which __dup_key were merged

    grouped = df_work.groupby("__dup_key", dropna=False)

    for key, group in grouped:
        if group["__skip_merge"].any():
        # Skip merging for this group
            for i in group.index:
                row = df.loc[i].copy()
                row[merged_flag_col] = "N"
                merged_rows.append(row)
            continue  # move to next group
        group_size = len(group)
        
        if group_size == 1:
            row = df.loc[group.index[0]].copy()
            row[merged_flag_col] = "N"
            merged_rows.append(row)
            continue

        # Determine if any rows have COMBINED INDIVIDUALS
        has_combined = False
        if "Contact Type" in df.columns:
            has_combined = any(group["Contact Type"].astype(str).str.upper() == "COMBINED INDIVIDUALS")

        # Merge rule:
        # - If no COMBINED INDIVIDUALS, merge normally
        # - If COMBINED INDIVIDUALS present AND group_size >= 4, merge
        # - Otherwise, keep rows as-is
        if has_combined:
            if group_size < 4:
                # Too small, do not merge
                for i in group.index:
                    row = df.loc[i].copy()
                    row[merged_flag_col] = "N"
                    merged_rows.append(row)
                continue

            # Step A: define unique row including First Name & numeric fields
            unique_key_cols = ["Owner (Standardized)", "First Name", "Address", "City", "State", "# of Interests", "County", "Target State"]
            group["__unique_row_key"] = group[unique_key_cols].astype(str).agg("|".join, axis=1)

            # Step B: count each unique row
            counts = group["__unique_row_key"].value_counts()
            eligible_keys = counts[counts > 1].index.tolist()  # repeat >=2

            # Step C: merge eligible keys
            for ukey in eligible_keys:
                subset = group[group["__unique_row_key"] == ukey]
                base = subset.iloc[0].copy()
                for field in merge_fields:
                    if field in df.columns:
                        total = subset[field].apply(safe_float).sum()
                        base[field] = total
                # Build "Owner Merge Info" column
                merged_owners = subset.apply(
                    lambda x: get_owner_with_details(
                        x["Owner (Standardized)"],
                        x.get("Owner ID", ""),
                        x.get("# of Interests", 0),
                        x.get("Total Value - Low ($)", 0)
                    ),
                    axis=1
                )
                base["Owner Merge Info"] = " | ".join(sorted(set(merged_owners)))  # unique and sorted
                base[merged_flag_col] = "Y"
                # Initialize Remarks
                base["Remarks"] = ""

                # Columns to check for identical values
                check_cols = [
                    "Owner (Standardized)", "Address", "City", "State", "# of Interests",
                    "Target State", "County", "Total Value - Low ($)", "PDP Value ($)", "Total Value - High ($)"
                ]

                # Apply remark only if Contact Type is NOT COMBINED INDIVIDUALS
                if "Contact Type" in df.columns and all(group["Contact Type"].astype(str).str.upper() != "COMBINED INDIVIDUALS"):
                    if group[check_cols].nunique().max() == 1:
                        base["Remarks"] = "Consider modifying CTT to COMBINED INDIVIDUALS"

                merged_rows.append(base)
                merged_groups_keys.add(key)



            # Step D: keep non-eligible rows as-is
            for i in group.index:
                if group.at[i, "__unique_row_key"] not in eligible_keys:
                    row = df.loc[i].copy()
                    row[merged_flag_col] = "N"
                    merged_rows.append(row)
            continue  # skip normal merging

        # ---- Normal merge for non-COMBINED INDIVIDUALS
        base = df.loc[group.index[0]].copy()
        for field in merge_fields:
            if field in df.columns:
                total = sum(safe_float(df.loc[i, field]) for i in group.index)
                base[field] = total
        # Build "Owner Merge Info" column
        merged_owners = group.apply(
            lambda x: get_owner_with_details(
                x["Owner (Standardized)"],
                x.get("Owner ID", ""),
                x.get("# of Interests", 0),
                x.get("Total Value - Low ($)", 0)
            ),
            axis=1
        )
        base["Owner Merge Info"] = " | ".join(sorted(set(merged_owners)))  # unique and sorted
        base[merged_flag_col] = "Y"
        # Initialize Remarks
        base["Remarks"] = ""

        # Check if all rows in group are identical in specified columns
        check_cols = [
            "Owner (Standardized)", "Address", "City", "State", "# of Interests",
            "Target State", "County", "Total Value - Low ($)", "PDP Value ($)", "Total Value - High ($)"
        ]

        # Only apply remark if Contact Type is NOT COMBINED INDIVIDUALS
        if "Contact Type" in df.columns and all(group["Contact Type"].astype(str).str.upper() != "COMBINED INDIVIDUALS"):
            if group[check_cols].nunique().max() == 1:
                base["Remarks"] = "Consider modifying CTT to COMBINED INDIVIDUALS"

        merged_rows.append(base)
        merged_groups_keys.add(key)


    merged_df = pd.DataFrame(merged_rows)

    # ---- NEW: Filter duplicate output to only merged rows
    duplicates_df = df_work[df_work["__dup_key"].isin(merged_groups_keys)].copy()
    duplicates_df.drop(columns=["__norm_owner", "__norm_address", "__norm_city",
                                "__norm_state", "__norm_county", "__norm_target_state",
                                "__norm_interests", "__dup_key", "__unique_row_key"],
                    inplace=True, errors="ignore")

    # ---- Step 5: Clean helper columns
    for col in ["__norm_owner", "__norm_address", "__norm_city", "__norm_state", "__norm_county",
                "__norm_target_state", "__norm_interests", "__dup_key", "__unique_row_key"]:
        if col in merged_df.columns:
            merged_df.drop(columns=[col], inplace=True, errors="ignore")

    # ---- Step 6: Prepare output filenames
    datetoday = datetime.now().strftime("%Y%m%d_%H%M%S")
    dup_file_name = f"duplicate_rows_{datetoday}_{file_name}"
    merged_file_name = f"merged_output_{datetoday}_{file_name}"


    # ---- Step 8: Save duplicate rows to output folder
    dup_path = os.path.join(output_folder, dup_file_name)
    duplicates_df.to_excel(dup_path, index=False)


    # ---- Step 9: Save merged output to output folder
    merged_path = os.path.join(output_folder, merged_file_name)
    merged_df.to_excel(merged_path, index=False)


    # ---- Step 10: Console summary
    summary_msg = (
        f"✅ Duplicate Check Complete for `{file_name}`\n\n"
        f"*Total rows:* {original_row_count}\n"
        f"*Duplicate rows detected:* {len(duplicates_df)}\n"
        f"*Final merged rows:* {len(merged_df)}\n\n"
        f"⚠️ _This output is generated automatically to assist in the review of duplicate owners and interests. "
        f"Please verify the results carefully before endorsing for BUDB loading, "
        f"as any errors may impact marketing and sales operations. "
        f"It is also recommended to check the corresponding WM file to ensure the # of Interests align._\n\n"
        f"_Please review the attached files._"
    )

    print(f"Processed {file_name}")
    print(f"  Total rows: {original_row_count}")
    print(f"  Duplicate rows: {len(duplicates_df)}")
    print(f"  Final merged rows: {len(merged_df)}")

if __name__ == "__main__":

    input_folder = "input"
    output_folder = "output"

    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Validate input folder
    if not os.path.isdir(input_folder):
        raise FileNotFoundError(f"Input folder not found: {input_folder}")

    files = [
        f for f in os.listdir(input_folder)
        if f.lower().endswith((".xlsx", ".xls"))
    ]

    if not files:
        print("No Excel files found in input folder.")
    else:
        print(f"Found {len(files)} file(s) to process.\n")

    for idx, file_name in enumerate(files, start=1):
        input_path = os.path.join(input_folder, file_name)

        print(f"[{idx}/{len(files)}] Processing: {file_name}")

        try:
            df = pd.read_excel(input_path)
            run_owner_interest_merger(df, file_name, output_folder)

        except Exception as e:
            print(f"❌ Failed to process {file_name}")
            print(f"   Reason: {e}")

        print("-" * 50)
