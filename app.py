"""
app.py
-------------------------
Streamlit web app for the Data Comparison Tool (Contact vs Salesforce).

This is a Streamlit port of comparator.py's core logic — matching,
deduplication, unidirectional/bidirectional comparison, and the
Campaign ID / AdGroup ID marketing filters, all producing the same
multi-sheet Excel report. The Tkinter GUI is replaced by Streamlit
widgets; the comparison logic itself is unchanged.

Run locally:
    streamlit run app.py
"""

import io
import re

import pandas as pd
import streamlit as st

# ── File Reader ─────────────────────────────────


def read_file(uploaded_file):
    name = uploaded_file.name.lower()
    try:
        if name.endswith(".csv"):
            try:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, encoding="utf-8-sig")
            except Exception:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, encoding="latin1")
        else:
            uploaded_file.seek(0)
            return pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        raise Exception(f"Error reading file: {uploaded_file.name}\n{e}")


# ── Normalization ─────────────────────────────────


def normalize_phone(p):
    if pd.isna(p):
        return None
    digits = re.sub(r"[^\d]", "", str(p))
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits if len(digits) >= 7 else None


def normalize_email(e):
    return str(e).strip().lower() if pd.notna(e) else ""


def normalize_name(first, last):
    return f"{str(first).strip()} {str(last).strip()}".strip()


# ── Deduplication ─────────────────────────────────


def deduplicate(df, use_phone, use_email, use_name):
    subset = []

    if use_phone and "_phone" in df:
        subset.append("_phone")
    if use_email and "_email" in df:
        subset.append("_email")
    if use_name and "_name" in df:
        subset.append("_name")

    if subset:
        df = df.drop_duplicates(subset=subset, keep="first")

    return df


# ── Match Key Builder ─────────────────────────────


def build_match_keys(df, use_phone, use_email, use_name):
    keys = []

    if use_phone:
        keys.append(df["_phone"].fillna(""))

    if use_email:
        keys.append(df["_email"].fillna(""))

    if use_name:
        keys.append(df["_name"].fillna(""))

    if not keys:
        raise Exception("Select at least one matching option")

    return pd.Series(list(zip(*keys)), index=df.index)


# ── Marketing Analysis ─────────────────────────────


def marketing_analysis(contact, sf, mode, direction, matched, excluded):
    for col in ["Campaign ID", "AdGroup ID"]:
        if col not in contact:
            contact[col] = ""
        if col not in sf:
            sf[col] = ""

    for df in [contact, sf]:
        df["Campaign ID"] = df["Campaign ID"].fillna("").astype(str).str.strip()
        df["AdGroup ID"] = df["AdGroup ID"].fillna("").astype(str).str.strip()

    contact_cmp = contact[["key", "Campaign ID", "AdGroup ID"]].copy()
    sf_cmp = sf[["key", "Campaign ID", "AdGroup ID"]].copy()

    contact_cmp.columns = ["Key", "Contact Campaign", "Contact AdGroup"]
    sf_cmp.columns = ["Key", "SF Campaign", "SF AdGroup"]

    contact_cmp = contact_cmp[contact_cmp["Key"].astype(str).str.strip() != ""]
    sf_cmp = sf_cmp[sf_cmp["Key"].astype(str).str.strip() != ""]

    contact_cmp = contact_cmp.drop_duplicates(subset=["Key"])
    sf_cmp = sf_cmp.drop_duplicates(subset=["Key"])

    merged = pd.merge(contact_cmp, sf_cmp, on="Key", how="outer")

    def campaign_status(row):
        c = row["Contact Campaign"]
        s = row["SF Campaign"]

        if c and s:
            return "Match" if c == s else "Different"
        elif c and not s:
            return "Missing in SF"
        elif s and not c:
            return "Missing in Contact"
        return "Missing Both"

    merged["Campaign Status"] = merged.apply(campaign_status, axis=1)

    def adgroup_status(row):
        c = row["Contact AdGroup"]
        s = row["SF AdGroup"]

        if c and s:
            return "Match" if c == s else "Different"
        elif c and not s:
            return "Missing in SF"
        elif s and not c:
            return "Missing in Contact"
        return "Missing Both"

    merged["AdGroup Status"] = merged.apply(adgroup_status, axis=1)

    contact_unique_keys = set(contact["key"])
    sf_unique_keys = set(sf["key"])
    total = len(contact_unique_keys.union(sf_unique_keys))

    marketing_summary = pd.DataFrame(
        {
            "Metric": [
                "Comparison Mode",
                "Direction",
                "Contact Records",
                "SF Records",
                "Total Unique Entities",
                "Matched Unique Entities",
                "Unmatched Unique Entities",
                "Matched Rows",
                "Excluded Rows",
                "Contact Missing Records",
                "SF Missing Records",
                "Match Percentage",
                "Missing Percentage",
            ],
            "Value": [
                mode,
                direction if mode == "UNIDIRECTIONAL" else "BOTH DIRECTIONS",
                len(contact),
                len(sf),
                total,
                len(set(contact["key"]).intersection(set(sf["key"]))),
                len(set(contact["key"]).symmetric_difference(set(sf["key"]))),
                len(matched),
                len(excluded),
                (
                    len(excluded[excluded.get("Missing From", "") == "Salesforce"])
                    if "Missing From" in excluded.columns
                    else 0
                ),
                (
                    len(excluded[excluded.get("Missing From", "") == "Contact"])
                    if "Missing From" in excluded.columns
                    else 0
                ),
                round(
                    (len(set(contact["key"]).intersection(set(sf["key"]))) / total) * 100,
                    2,
                )
                if total
                else 0,
                round(
                    (
                        len(set(contact["key"]).symmetric_difference(set(sf["key"])))
                        / total
                    )
                    * 100,
                    2,
                )
                if total
                else 0,
            ],
        }
    )

    return merged, marketing_summary


# ── Core Logic ────────────────────────────────────


def run_extraction(
    contact_files,
    sf_files,
    use_phone,
    use_email,
    use_name,
    filter_campaign,
    filter_adgroup,
    mode,
    direction,
    log,
):
    if not contact_files:
        raise Exception("Please upload Contact file(s)")

    if not sf_files:
        raise Exception("Please upload Salesforce file(s)")

    log("Loading Contact files...")
    contact = pd.concat([read_file(f) for f in contact_files], ignore_index=True)
    contact["Source"] = "Contact"

    log("Loading Salesforce files...")
    sf = pd.concat([read_file(f) for f in sf_files], ignore_index=True)
    sf["Source"] = "Salesforce"

    # Normalize
    if use_phone:
        contact["_phone"] = contact["Phone"].apply(normalize_phone)
        sf["_phone"] = sf["Phone"].apply(normalize_phone)

    if use_email:
        contact["_email"] = contact["Email"].apply(normalize_email)
        sf["_email"] = sf["Email"].apply(normalize_email)

    if use_name:
        contact["_name"] = contact.apply(
            lambda x: normalize_name(x.get("First Name", ""), x.get("Last Name", "")),
            axis=1,
        )
        sf["_name"] = sf.apply(
            lambda x: normalize_name(x.get("First Name", ""), x.get("Last Name", "")),
            axis=1,
        )

    # Deduplication
    contact = deduplicate(contact, use_phone, use_email, use_name)
    sf = deduplicate(sf, use_phone, use_email, use_name)

    log(f"After dedup → Contact: {len(contact)}, SF: {len(sf)}")

    # Build match keys
    contact["key"] = build_match_keys(contact, use_phone, use_email, use_name)
    sf["key"] = build_match_keys(sf, use_phone, use_email, use_name)

    contact_keys = set(contact["key"])
    sf_keys = set(sf["key"])

    # ── Comparison Modes ─────────────────────

    if mode == "UNIDIRECTIONAL":
        if direction == "CONTACT_TO_SF":
            contact["Match"] = contact["key"].isin(sf_keys)
            matched = contact[contact["Match"]]
            excluded = contact[~contact["Match"]]
        else:
            sf["Match"] = sf["key"].isin(contact_keys)
            matched = sf[sf["Match"]]
            excluded = sf[~sf["Match"]]

    else:  # BIDIRECTIONAL
        contact["Match"] = contact["key"].isin(sf_keys)
        sf["Match"] = sf["key"].isin(contact_keys)

        contact_missing = contact[~contact["Match"]].copy()
        contact_missing["Missing From"] = "Salesforce"

        sf_missing = sf[~sf["Match"]].copy()
        sf_missing["Missing From"] = "Contact"

        matched = pd.concat([contact[contact["Match"]], sf[sf["Match"]]])
        excluded = pd.concat([contact_missing, sf_missing])

    # ── Optional Marketing Filters ─────────────────

    for col in ["Campaign ID", "AdGroup ID"]:
        if col not in contact:
            contact[col] = ""
        if col not in sf:
            sf[col] = ""

        contact[col] = contact[col].fillna("").astype(str).str.strip()
        sf[col] = sf[col].fillna("").astype(str).str.strip()

    if filter_campaign:
        contact_campaigns = set(contact["Campaign ID"][contact["Campaign ID"] != ""])
        sf_campaigns = set(sf["Campaign ID"][sf["Campaign ID"] != ""])

        if direction == "SF_TO_CONTACT" or mode == "BIDIRECTIONAL":
            sf_missing_campaign = sf[
                (~sf["Campaign ID"].isin(contact_campaigns)) & (sf["Campaign ID"] != "")
            ].copy()
            sf_missing_campaign["Exclusion Reason"] = "Campaign ID Missing in Contact"
            excluded = pd.concat([excluded, sf_missing_campaign], ignore_index=True)

        if direction == "CONTACT_TO_SF" or mode == "BIDIRECTIONAL":
            contact_missing_campaign = contact[
                (~contact["Campaign ID"].isin(sf_campaigns)) & (contact["Campaign ID"] != "")
            ].copy()
            contact_missing_campaign["Exclusion Reason"] = "Campaign ID Missing in SF"
            excluded = pd.concat([excluded, contact_missing_campaign], ignore_index=True)

    if filter_adgroup:
        contact_adgroups = set(contact["AdGroup ID"][contact["AdGroup ID"] != ""])
        sf_adgroups = set(sf["AdGroup ID"][sf["AdGroup ID"] != ""])

        if direction == "SF_TO_CONTACT" or mode == "BIDIRECTIONAL":
            sf_missing_adgroup = sf[
                (~sf["AdGroup ID"].isin(contact_adgroups)) & (sf["AdGroup ID"] != "")
            ].copy()
            sf_missing_adgroup["Exclusion Reason"] = "AdGroup ID Missing in Contact"
            excluded = pd.concat([excluded, sf_missing_adgroup], ignore_index=True)

        if direction == "CONTACT_TO_SF" or mode == "BIDIRECTIONAL":
            contact_missing_adgroup = contact[
                (~contact["AdGroup ID"].isin(sf_adgroups)) & (contact["AdGroup ID"] != "")
            ].copy()
            contact_missing_adgroup["Exclusion Reason"] = "AdGroup ID Missing in SF"
            excluded = pd.concat([excluded, contact_missing_adgroup], ignore_index=True)

    excluded = excluded.drop_duplicates(subset=["key"])

    campaign_detail_df, marketing_df = marketing_analysis(
        contact, sf, mode, direction, matched, excluded
    )

    # ── Summary ─────────────────────────────

    contact_rows = len(contact)
    sf_rows = len(sf)
    total_physical_rows = contact_rows + sf_rows

    unique_total_keys = len(set(contact["key"]).union(set(sf["key"])))
    unique_matched_keys = len(set(contact["key"]).intersection(set(sf["key"])))
    unique_excluded_keys = unique_total_keys - unique_matched_keys

    contact_matched_rows = contact["Match"].sum() if "Match" in contact else 0
    sf_matched_rows = sf["Match"].sum() if "Match" in sf else 0
    total_matched_rows = contact_matched_rows + sf_matched_rows

    summary = pd.DataFrame(
        {
            "Metric": [
                "Mode",
                "Direction",
                "Contact Rows",
                "SF Rows",
                "Total Physical Rows",
                "Unique Combined Keys",
                "Unique Matched Keys",
                "Unique Excluded Keys",
                "Contact Matched Rows",
                "SF Matched Rows",
                "Total Matched Rows",
                "Unique Match %",
                "Unique Missing %",
            ],
            "Value": [
                mode,
                direction if mode == "UNIDIRECTIONAL" else "Two-Way Comparison",
                contact_rows,
                sf_rows,
                total_physical_rows,
                unique_total_keys,
                unique_matched_keys,
                unique_excluded_keys,
                contact_matched_rows,
                sf_matched_rows,
                total_matched_rows,
                round((unique_matched_keys / unique_total_keys) * 100, 2)
                if unique_total_keys
                else 0,
                round((unique_excluded_keys / unique_total_keys) * 100, 2)
                if unique_total_keys
                else 0,
            ],
        }
    )

    # ── Build Excel in memory ─────────────────────────

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        contact.to_excel(writer, sheet_name="Contact Data", index=False)
        sf.to_excel(writer, sheet_name="SF Data", index=False)
        matched.to_excel(writer, sheet_name="Matched", index=False)
        excluded.to_excel(writer, sheet_name="Excluded", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
        campaign_detail_df.to_excel(writer, sheet_name="Campaign Comparison", index=False)
        marketing_df.to_excel(writer, sheet_name="Marketing Summary", index=False)
    buffer.seek(0)

    log("══════════════════════════════")
    log(f"MODE: {mode}")

    matching_used = []
    if use_phone:
        matching_used.append("Phone")
    if use_email:
        matching_used.append("Email")
    if use_name:
        matching_used.append("Name")

    log(f"Matching Using: {', '.join(matching_used)}")
    log(f"Contact Records: {len(contact)}")
    log(f"SF Records: {len(sf)}")
    log(f"Unique Matched Entities: {unique_matched_keys}")
    log(f"Unique Excluded Entities: {unique_excluded_keys}")
    log(f"Matched Rows Sheet Count: {len(matched)}")
    log(f"Excluded Rows Sheet Count: {len(excluded)}")

    try:
        campaign_match = (campaign_detail_df["Campaign Status"] == "Match").sum()
        campaign_missing_sf = (campaign_detail_df["Campaign Status"] == "Missing in SF").sum()
        campaign_missing_contact = (
            campaign_detail_df["Campaign Status"] == "Missing in Contact"
        ).sum()

        log(f"Campaign Matches: {campaign_match}")
        log(f"Campaign Missing in SF: {campaign_missing_sf}")
        log(f"Campaign Missing in Contact: {campaign_missing_contact}")
    except Exception:
        pass

    log("Output Ready")
    log("══════════════════════════════")

    return {
        "buffer": buffer,
        "summary": summary,
        "matched": matched,
        "excluded": excluded,
        "unique_matched_keys": unique_matched_keys,
        "unique_excluded_keys": unique_excluded_keys,
    }


# ── Streamlit UI ────────────────────────────────────

st.set_page_config(page_title="Data Comparison Tool", page_icon="📊", layout="wide")

st.title("📊 Data Comparison Tool")
st.caption("Compare Contact files against Salesforce data by Phone, Email, and/or Name.")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📂 Contact Files")
    contact_files = st.file_uploader(
        "Upload Contact file(s)",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
        key="contact_files",
    )

with col2:
    st.subheader("📂 Salesforce Files")
    sf_files = st.file_uploader(
        "Upload Salesforce file(s)",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
        key="sf_files",
    )

st.subheader("⚙️ Matching Options")
m1, m2, m3 = st.columns(3)
use_phone = m1.checkbox("Phone", value=True)
use_email = m2.checkbox("Email", value=True)
use_name = m3.checkbox("Name", value=False)

st.subheader("📢 Marketing Filters")
f1, f2 = st.columns(2)
filter_campaign = f1.checkbox("Exclude Missing Campaign IDs", value=False)
filter_adgroup = f2.checkbox("Exclude Missing AdGroup IDs", value=False)

st.subheader("🔄 Comparison Mode")
mode_label = st.radio("Mode", ["Unidirectional", "Bidirectional"], horizontal=True)
mode = "UNIDIRECTIONAL" if mode_label == "Unidirectional" else "BIDIRECTIONAL"

direction = "CONTACT_TO_SF"
if mode == "UNIDIRECTIONAL":
    direction_label = st.radio(
        "Direction", ["Contact → SF", "SF → Contact"], horizontal=True
    )
    direction = "CONTACT_TO_SF" if direction_label == "Contact → SF" else "SF_TO_CONTACT"

run_clicked = st.button("🚀 Run Comparison", type="primary")

if run_clicked:
    logs = []

    def log(msg):
        logs.append(msg)

    try:
        with st.spinner("Running comparison..."):
            result = run_extraction(
                contact_files,
                sf_files,
                use_phone,
                use_email,
                use_name,
                filter_campaign,
                filter_adgroup,
                mode,
                direction,
                log,
            )
    except Exception as e:
        st.error(f"❌ {e}")
        if logs:
            with st.expander("Log"):
                st.text("\n".join(logs))
    else:
        st.success("Comparison complete!")

        c1, c2, c3 = st.columns(3)
        c1.metric("Unique Matched", result["unique_matched_keys"])
        c2.metric("Unique Excluded", result["unique_excluded_keys"])
        c3.metric("Excluded Rows", len(result["excluded"]))

        st.subheader("Summary")
        st.dataframe(result["summary"], use_container_width=True, hide_index=True)

        st.download_button(
            label="💾 Download Full Report (Excel)",
            data=result["buffer"],
            file_name="comparison_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("Log"):
            st.text("\n".join(logs))

        with st.expander("Preview: Excluded (first 100 rows)"):
            st.dataframe(result["excluded"].head(100), use_container_width=True)
