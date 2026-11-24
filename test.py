import sys
import pandas as pd
import os
from datetime import datetime, date
import warnings
warnings.filterwarnings('ignore')
import locale
locale.setlocale(locale.LC_ALL, 'en_IN.utf8')
import pyodbc
from sqlalchemy import create_engine



def get_today_files(base_folder):
    from datetime import datetime, date
    import os

    today = date.today()
    valid_files = []

    # -------------------------------
    # TARGET MONTH FOLDER (Apr'26)
    # -------------------------------
    folder_name = datetime.today().strftime("%b'%y")
    target_folder = os.path.join(base_folder, folder_name)
    # print(folder_name)
    # print(target_folder)


    if not os.path.exists(target_folder):
        print(f"Folder not found: {target_folder}")
        return []

    # -------------------------------
    # ONLY CHANGE HERE (no os.walk)
    # -------------------------------

    print(os.listdir(target_folder))
    for file in os.listdir(target_folder):
        if file.endswith((".xlsx", ".xlsb")):
            file_path = os.path.join(target_folder, file)
            # print(file_path)

            try:
                
                modified_time = datetime.fromtimestamp(
                    os.path.getmtime(file_path)
                ).date()
                # print(modified_time)
                if modified_time == today:
                    valid_files.append(file_path)

            except Exception as e:
                print(f"Skipping file: {file_path}")
                print(f"Actual error: {e}")

    return valid_files




# if  valid_files:
#     print("files found today. Exiting.")
# else:
#     print('no files there')    
# print(valid_files)python test.py
# for x in valid_files:
#     print(x)




# -------------------------------
def ingest_files(valid_files, log_path, mapping_path):

    dfs = []
    log_data = []
    today_str = datetime.today().strftime("%Y-%m-%d")

    # -------------------------------
    # REQUIRED SCHEMA
    # -------------------------------
    correct_cols = [
        "Leadid","BookingDate","FY_Year","BookingMonth",
        "IssuanceDate","RiskStartDate","RiskEndDate",
        "PolicyNo","InsurerName","BU","Product","SubProduct",
        "VehicleSubClass","PlanType","PlanName","PolicyStatus",
        "ODPremium","TPPremium","NetPremium","TotalPremium",
        "PointOn","Points","IRDA","Reward","OtherDues","FinalPayin",
        "Remarks","Status","ReceivedMonth","FirstYrAmt","FifthYrAmt",
        "Additional_Rec","Total_Rec","Diff_Amt",
        "RewardOD","RewardTP","IncentiveOD","IncentiveTP"
    ]

    date_cols = ["BookingDate","IssuanceDate","RiskStartDate","RiskEndDate"]

    # -------------------------------
    # GET FILES
    # -------------------------------
    file_list = valid_files

    print(f" Total today's files: {len(file_list)}")

    # -------------------------------
    # LOOP FILES
    # -------------------------------
    for file_path in file_list:

        file = os.path.basename(file_path)
        print()
        print(f"Reading: {file_path}")

        status = "Success"
        error_msg = None
        row_count = 0

        try:
            # -------------------------------
            # READ FILE
            # -------------------------------
            if file.endswith(".xlsb"):
                temp_df = pd.read_excel(
                    file_path,
                    engine="pyxlsb",
                    sheet_name=0,
                    keep_default_na=False
                )

                # XLSB DATE FIX
                for col in date_cols:
                    if col in temp_df.columns:
                        temp_df[col] = temp_df[col].apply(
                            lambda x: pd.to_datetime(
                                x, origin="1899-12-30", unit="D", errors="coerce"
                            ) if isinstance(x, (int, float))
                            else pd.to_datetime(x, errors="coerce")
                        )

            else:
                temp_df = pd.read_excel(
                    file_path,
                    engine="openpyxl",
                    sheet_name=0,
                    keep_default_na=False
                )
                
                for col in date_cols:
                    if col in temp_df.columns:
                        temp_df[col] = temp_df[col].apply(
                            lambda x: pd.to_datetime(
                                x, origin="1899-12-30", unit="D", errors="coerce"
                            ) if isinstance(x, (int, float))
                            else pd.to_datetime(x, errors="coerce")
                        )

            

            # -------------------------------
            # CLEAN COLUMNS
            # -------------------------------
            temp_df.columns = temp_df.columns.str.strip()

            temp_df = temp_df.loc[
            :,
            ~temp_df.columns.str.contains("^Unnamed", case=False, na=False)
            ]

            # -------------------------------
            # SCHEMA VALIDATION
            # -------------------------------
            file_cols = {
                            col.strip().lower(): col
                            for col in temp_df.columns
                        }
            
            # Expected columns: lowercase version -> standard column name
            expected_cols = {
                col.strip().lower(): col
                for col in correct_cols
                             }

            missing_cols = list(set(expected_cols) - set(file_cols))
            extra_cols = list(set(file_cols) - set(expected_cols))

            if missing_cols or extra_cols:
                status = "Failed"
                error_msg = f"Missing: {missing_cols} | Extra: {extra_cols}"

            else:
                # Convert whatever casing is present
                # into your standard column names
                rename_mapping = {
                    file_cols[key]: expected_cols[key]
                    for key in expected_cols
                }

                print(rename_mapping)
            
                temp_df = temp_df.rename(columns=rename_mapping)
            
                # Enforce standard column order
                temp_df = temp_df[correct_cols]
            
                row_count = len(temp_df)
            
                dfs.append(temp_df)

        except Exception as e:
            status = "Failed"
            error_msg = str(e)

        # -------------------------------
        # LOG ENTRY (ALWAYS LOG)
        # -------------------------------
        log_data.append({
            "Date": today_str,
            "FileName": file,
            "FilePath": file_path,
            "Status": status,
            "Rows": row_count,
            "Error": error_msg
        })

    # -------------------------------
    # FINAL DF
    # -------------------------------
    final_df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

    print(f" Final rows appended: {len(final_df)}")

    # -------------------------------
    # SAVE LOG (APPEND MODE)
    # -------------------------------
    log_df = pd.DataFrame(log_data)

    try:
        if os.path.exists(log_path):
            old_log = pd.read_excel(log_path)
            log_df = pd.concat([old_log, log_df], ignore_index=True)

        log_df.to_excel(log_path, index=False)
        print(f" Log saved: {log_path}")

    except Exception as e:
        print(" Log file is open/locked, skipping write")

    # -------------------------------
    # LOAD MAPPING FILE
    # -------------------------------
    df_insurername = pd.read_excel(
        mapping_path,
        engine="openpyxl",
        keep_default_na=False
    )

    return final_df, df_insurername, log_df



def map_insurer_names(df, df_insurername):
    from rapidfuzz import process, fuzz, utils
    import pandas as pd

    # -------------------------------
    # Mapping table (lowercase key used for every comparison)
    # -------------------------------
    mapping_df = df_insurername[["InsurerName", "new insurer name"]].copy().reset_index(drop=True)
    mapping_df["key"] = mapping_df["InsurerName"].astype(str).str.strip().str.lower()

    choices = mapping_df["key"].tolist()          # same order/index as mapping_df
    choices_set = set(choices)
    print(df["InsurerName"].drop_duplicates().tolist())
    # -------------------------------
    # Step 1: find unmatched names (case-insensitive exact check)
    # -------------------------------
    df_key = df["InsurerName"].astype(str).str.strip().str.lower()
    

    unmatched = (
        df.loc[~df_key.isin(choices_set), "InsurerName"]
        .dropna()
        .drop_duplicates()
    )

    # -------------------------------
    # Step 2: fuzzy match function
    # -------------------------------
    def get_best_match(name):
        name = str(name).strip().lower()           # force lowercase on the query too
        match = process.extractOne(
            name,
            choices,
            scorer=fuzz.token_set_ratio,
            processor=utils.default_process,       # lowercases + cleans both sides
            score_cutoff=80
        )
        if match:
            choice, score, idx = match
            new_name = mapping_df["new insurer name"].iloc[idx]
            return choice, new_name, score
        return None, None, None

    # -------------------------------
    # Step 3: build result
    # -------------------------------
    results = []
    for name in unmatched:
        matched_name, new_name, score = get_best_match(name)
        results.append({
            "Original_Name": name,
            "Matched_Name": matched_name,
            "Match_Score": score,
            "New_Insurer_Name": new_name
        })

    fuzzy_df = pd.DataFrame(results)

    if fuzzy_df.empty:
        print("No unmatched insurers found, nothing to add")
        new_rows = pd.DataFrame(columns=["InsurerName", "new insurer name"])
    else:
        print(fuzzy_df)                            # review what was matched
        new_rows = fuzzy_df[["Original_Name", "New_Insurer_Name"]].dropna()
        new_rows.columns = ["InsurerName", "new insurer name"]

    # only add names not already present (case-insensitive)
    existing_keys = set(mapping_df["key"])
    new_rows_filtered = new_rows[
        ~new_rows["InsurerName"].astype(str).str.strip().str.lower().isin(existing_keys)
    ]

    # -------------------------------
    # Append safely
    # -------------------------------
    updated_mapping = pd.concat(
        [df_insurername, new_rows_filtered],
        ignore_index=True
    )

    # -------------------------------
    # Save mapping
    # -------------------------------
    output_path = r"\\10.0.11.37\bd_pbp\DB Upload\Revenue_file_ingestion\Insurer name Vs db_updated.xlsx"
    updated_mapping.to_excel(output_path, index=False)

    # -------------------------------
    # Apply mapping (lowercase keys, so case differences never break it)
    # -------------------------------
    mapping_dict = dict(
        zip(
            updated_mapping["InsurerName"].astype(str).str.strip().str.lower(),
            updated_mapping["new insurer name"]
        )
    )

    df["InsurerName"] = (
        df["InsurerName"].astype(str).str.strip().str.lower().map(mapping_dict)
    )

    print(df["InsurerName"].unique())
    print(df[df["InsurerName"].isna()])

    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    return df, updated_mapping
     




base_folder = r"\\10.0.11.37\bd_pbp\DB Upload\Revenue_cal_table Update"
log_path = r"\\10.0.11.37\bd_pbp\DB Upload\Revenue_file_ingestion\ingestion_log.xlsx"
mapping_path = r"\\10.0.11.37\bd_pbp\DB Upload\Revenue_file_ingestion\Insurer name Vs db_updated.xlsx"

valid_files=get_today_files(base_folder)

# if valid_files:

# print()
# print(valid_files)


if not valid_files:
    print("No files found today. Exiting.")
    sys.exit(0)


df, df_insurername, log_df = ingest_files(valid_files, log_path, mapping_path)

df, updated_mapping=map_insurer_names(df, df_insurername)

print(df['InsurerName'].head())
# print(updated_mapping)

# print(log_df)


