import pandas as pd
from io import BytesIO

def categorize_benefit(benefit_text):
    if not isinstance(benefit_text, str):
        return "No Match"
    text = benefit_text.upper()

    if any(x in text for x in [
        "OUT PATIENT OVERALL",
        "ANTE AND POST NATAL CARE",
        "IMMUNIZATION",
        "HEALTH CHECKUP",
        "WELLBEING BENEFIT",
        "COPAY KES 1000",
        "COPAY 1 TIER"
    ]):
        return "OUTPATIENT"
    elif any(x in text for x in [
        "CONGENITAL", "CHILDBRITH", "NEO NATAL", "PREMATURITY",
        "EXTERNAL MEDICAL APPLIANCES",
        "NON ACCIDENTAL DENTAL",
        "NON ACCIDENTAL OPTICAL",
        "HOSPITALIZATION",
        "PRE-EXISTING", "CHRONIC",
        "PSYCHIATRY", "PSYCHOTHERAPY",
        "POST HOSPITALIZATION"
    ]):
        return "INPATIENT"
    elif "DENTAL" in text:
        return "DENTAL"
    elif any(x in text for x in ["OPTICAL", "FRAMES"]):
        return "OPTICAL"
    elif "LAST EXPENSE" in text:
        return "LAST EXPENSE"
    elif any(x in text for x in ["NORMAL DELIVERY", "EMERGENCY CEASEREAN"]):
        return "MATERNITY"
    else:
        return "No Match"

def run_preprocessing(file):
    # Read "Export" sheet from uploaded in-memory Excel file
    df = pd.read_excel(file, sheet_name="Export", engine="openpyxl")
    last_row = df.shape[0]

    # BENEFIT column (after column O, which is index 14)
    df.insert(15, "BENEFIT", df.iloc[:, 14].apply(categorize_benefit))

    # MEMBER NO + TRANS DATE (E = 4, S = 18)
    df.insert(43, "MEMBER + TRANS DATE", df.iloc[:, 4].astype(str) + df.iloc[:, 18].astype(str))
    df.sort_values(by=df.columns[43], inplace=True, ignore_index=True)

    # COUNT column
    count_flags = [1]
    for i in range(1, last_row):
        count_flags.append(1 if df.iloc[i, 43] != df.iloc[i - 1, 43] else 0)
    df.insert(44, "COUNT", count_flags)

    # Load provider mapping from local CSV
    # provider_df = pd.read_csv("provider.csv")
    # provider_dict = dict(zip(provider_df.iloc[:, 0], provider_df.iloc[:, 1]))

    # Add PARENT PROVIDER after V (index 21)
    # df.insert(22, "PARENT PROVIDER", df.iloc[:, 21].map(provider_dict))

    # Sort by MEMBER NO (E = index 4)
    df.sort_values(by=df.columns[4], inplace=True, ignore_index=True)

    # UNIQUE COUNT column
    unique_flags = [1]
    for i in range(1, last_row):
        unique_flags.append(1 if df.iloc[i, 4] != df.iloc[i - 1, 4] else 0)
    df.insert(5, "UNIQUE COUNT", unique_flags)

    return df