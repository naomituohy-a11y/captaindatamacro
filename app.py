import json
import re
from io import BytesIO

import pandas as pd
import streamlit as st
from rapidfuzz.fuzz import ratio


def safe_json_loads(x):
    if pd.isna(x) or x == "" or x is None:
        return None
    if isinstance(x, (dict, list)):
        return x
    try:
        return json.loads(x)
    except Exception:
        return None


def fix_job_period_key(experiences):
    """Captain Data sometimes uses job_period; macro replaces with job_time_period."""
    if not isinstance(experiences, list):
        return experiences
    fixed = []
    for item in experiences:
        if isinstance(item, dict):
            if "job_period" in item and "job_time_period" not in item:
                item = dict(item)
                item["job_time_period"] = item.pop("job_period")
        fixed.append(item)
    return fixed


def pick_current_experience(experiences):
    """
    Choose current experience:
    - Prefer one with 'Present' in job_time_period (or job_period if still present)
    - Else first experience
    """
    if not isinstance(experiences, list) or len(experiences) == 0:
        return None

    def get_period(e):
        if not isinstance(e, dict):
            return ""
        return str(e.get("job_time_period") or e.get("job_period") or "")

    # Prefer 'Present'
    for e in experiences:
        if "present" in get_period(e).lower():
            return e

    return experiences[0]


def years2_invalid(period_text: str, years_list):
    """
    Mimics your Excel logic:
    Valid unless the period contains any token from Years2.
    (Your formula: if SEARCH(Years2, text) finds anything -> Invalid)
    """
    if period_text is None:
        return ""
    s = str(period_text).strip()
    if s == "":
        return ""
    for y in years_list:
        y = y.strip()
        if not y:
            continue
        if y.lower() in s.lower():
            return "Invalid"
    return "Valid"


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="processed")
    return out.getvalue()


st.set_page_config(page_title="Captain Data Processor", layout="wide")
st.title("Captain Data – New Export Processor (37-column format)")

uploaded = st.file_uploader("Upload NEW Captain Data CSV", type=["csv"])
output_format = st.selectbox("Output format", ["CSV", "XLSX"])

st.subheader("Optional: Years2 invalidation list (like your macro)")
years2_text = st.text_area(
    "Enter tokens (one per line). If a job_time_period contains any of these tokens, mark Invalid.",
    value="2018\n2017\n2016\n2015\n2014\n2013\n2012\n2011\n2010\n2009\n2008\n2007\n2006\n2005\n2004\n2003\n2002\n2001\n2000",
    height=180,
)
years2_list = [line.strip() for line in years2_text.splitlines() if line.strip()]

if uploaded:
    df = pd.read_csv(uploaded)

    # Parse JSON columns that matter
    df["experiences_json"] = df["experiences"].apply(safe_json_loads)
    df["experiences_json"] = df["experiences_json"].apply(fix_job_period_key)

    # Derived fields (keeps your new 37 cols intact; we append extras)
    current_period = []
    current_company = []
    current_title = []
    current_li_company_url = []

    for exps in df["experiences_json"]:
        cur = pick_current_experience(exps)
        if isinstance(cur, dict):
            current_period.append(cur.get("job_time_period") or cur.get("job_period") or "")
            current_company.append(cur.get("company_name") or "")
            current_title.append(cur.get("title") or cur.get("job_title") or "")
            current_li_company_url.append(cur.get("linkedin_company_url") or cur.get("linkedin_company_url_cleaned") or "")
        else:
            current_period.append("")
            current_company.append("")
            current_title.append("")
            current_li_company_url.append("")

    df["current_job_time_period"] = current_period
    df["current_company_name_from_exp"] = current_company
    df["current_job_title_from_exp"] = current_title
    df["current_linkedin_company_url_from_exp"] = current_li_company_url

    # Years2 validation (like Macro1/Macro7)
    df["current_job_time_period_validity"] = df["current_job_time_period"].apply(
        lambda s: years2_invalid(s, years2_list)
    )

    # Optional: quick similarity checks (useful replacement for String_Similarity)
    # (Example: compare company_name vs current_company_name_from_exp)
    df["company_name_similarity_to_exp"] = df.apply(
        lambda r: (ratio(str(r.get("company_name", "")), str(r.get("current_company_name_from_exp", ""))) / 100.0)
        if str(r.get("company_name", "")).strip() and str(r.get("current_company_name_from_exp", "")).strip()
        else "",
        axis=1
    )

    st.success(f"Processed {len(df):,} rows.")
    st.dataframe(df.head(20), use_container_width=True)

    if output_format == "CSV":
        csv_bytes = df.to_csv(index=False).encode("utf-8")
        st.download_button("Download processed CSV", data=csv_bytes, file_name="captain_data_processed.csv")
    else:
        xlsx_bytes = to_excel_bytes(df)
        st.download_button("Download processed XLSX", data=xlsx_bytes, file_name="captain_data_processed.xlsx")
