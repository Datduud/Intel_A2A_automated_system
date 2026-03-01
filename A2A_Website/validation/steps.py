import pandas as pd
import os

def list_all_step_functions():
    import sys
    import types
    current_module = sys.modules[__name__]
    return [
        name for name, obj in current_module.__dict__.items()
        if callable(obj) and name.endswith('_step8_save') or name.startswith('china_step') or name.startswith('india_step')
    ]
    
def china_step1_load_and_clean(input_path, year, month, **kwargs):
    df = pd.read_excel(input_path)
    df.columns = [col.strip() for col in df.columns]
    cols_needed = [
        "Freight Supplier", "HAWB / Bill of Lading #", "Custom Form Declaration Date",
        "Origin Airport/Port", "Item Description (English)", "Inco Term",
        "Supplier Name", "Number of Cartons"
    ]
    df = df[cols_needed]
    df["HAWB / Bill of Lading #"] = df["HAWB / Bill of Lading #"].astype(str).str.strip()
    df["Number of Cartons"] = pd.to_numeric(df["Number of Cartons"], errors='coerce').fillna(0).astype(int)
    df["Custom Form Declaration Date"] = pd.to_datetime(df["Custom Form Declaration Date"], errors='coerce')
    df["Year"] = df["Custom Form Declaration Date"].dt.year
    df["Month"] = df["Custom Form Declaration Date"].dt.month
    df = df[(df["Year"] == year) & (df["Month"] == month)]
    return df

def china_step2_unique_hawb(df, **kwargs):
    df = df.sort_values("HAWB / Bill of Lading #").reset_index(drop=True)
    df["Unique Order"] = ~df["HAWB / Bill of Lading #"].duplicated(keep='first')
    df["Unique Order"] = df["Unique Order"].astype(int)
    df["Total Order"] = 1
    df = df[df["Unique Order"] == 1].copy()
    return df

def china_step3_wd_tagging(df, **kwargs):
    wd_keywords = {"DIE", "HWS CARRIER"}
    df["WD_Tag"] = df["Item Description (English)"].str.upper().str.strip().apply(
        lambda x: "WD" if x in wd_keywords else None
    )
    wd_hawb = df[df["WD_Tag"] == "WD"]["HAWB / Bill of Lading #"].unique()
    df = df[~df["HAWB / Bill of Lading #"].isin(wd_hawb)].copy()
    df.drop(columns=["WD_Tag"], inplace=True)
    return df

def china_step4_fg_merge(df, hawb_path, **kwargs):
    fg_df = pd.read_excel(hawb_path, skiprows=2)
    fg_df.columns = fg_df.columns.str.strip()
    fg_df = fg_df[["FG_WD_CAPITAL[ORDER_RELEASE_XID]", "FG_WD_CAPITAL[COMMODITY]"]]
    fg_df = fg_df[fg_df["FG_WD_CAPITAL[COMMODITY]"].str.upper().str.strip() == "FG"]
    fg_keys = fg_df["FG_WD_CAPITAL[ORDER_RELEASE_XID]"].dropna().unique()
    df["FG_Match"] = df["HAWB / Bill of Lading #"].apply(lambda x: "Yes" if x in fg_keys else None)
    return df

def china_step5_remarks(df, **kwargs):
    air_freight_list = {"SHA", "PVG", "CTU", "CGO", "SZX", "PEK", "BJS", "CKG"}
    df["Remarks"] = df["Origin Airport/Port"].str.strip().str.upper().apply(
        lambda x: "Internal Trucking" if x in air_freight_list else None
    )

    def final_remark(row):
        if not row["Remarks"]:
            inco = str(row["Inco Term"]).strip().upper()
            supplier = str(row["Supplier Name"]).strip().upper()
            if inco not in {"FCA", "EXW"} and "INTEL" not in supplier:
                return "Supplier Paid"
        return row["Remarks"]

    df["Remarks"] = df.apply(final_remark, axis=1)
    return df

def china_step6_pbi_merge(df, hawb_path, **kwargs):
    nfg_df = pd.read_excel(hawb_path, sheet_name="IntelOpsReportMainALL", skiprows=2)
    nfg_df.columns = nfg_df.columns.str.strip()
    nfg_df["NFG_List[HAWB/BOL]"] = nfg_df["NFG_List[HAWB/BOL]"].astype(str).str.strip().str.upper()
    df["HAWB_Upper"] = df["HAWB / Bill of Lading #"].astype(str).str.strip().str.upper()
    df["In PBI"] = df["HAWB_Upper"].isin(nfg_df["NFG_List[HAWB/BOL]"]).map({True: "Yes", False: "No"})
    return df

def china_step7_kpi(df, **kwargs):
    df["PBI Count"] = (df["In PBI"] == "Yes").astype(int)
    df["Failed Order"] = df["Total Order"] - df["PBI Count"]
    df["Ontime %"] = df.apply(
        lambda row: round(row["PBI Count"] * 100 / row["Total Order"], 2)
        if row["Total Order"] else None,
        axis=1
    )
    return df

def china_step8_save(df, output_folder, **kwargs):
    output_path = os.path.join(output_folder, "validated_result.xlsx")
    df.to_excel(output_path, index=False)
    return output_path

# Example: If you want to add "india_step1_load_and_clean", "vietnam_step2_unique_hawb" etc,
# just copy one of the above functions and change the function name and, if needed, any logic inside.
