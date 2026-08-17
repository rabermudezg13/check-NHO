import io
import re
import unicodedata
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


st.set_page_config(page_title="NHO Attendance Checker", page_icon="✅", layout="wide")

FIELD_ALIASES = {
    "Applicant Name": ["applicant name", "talent name"],
    "Talent Email": ["talent email", "email"],
    "Current Onboarding Stage": ["current onboarding stage"],
    "HP Overview Template": ["hp overview template"],
    "OB365": ["ob365"],
    "District Release": ["district release"],
    "Pre-hire Training": ["pre hire training", "prehire training", "classmarts", "classmarts cert"],
    "Safety & Security": ["safety security", "policy explainer cert"],
    "Onboarding Introduction Email": ["onboarding introduction email"],
    "I-9": ["i 9", "i9"],
    "Education Requirement": ["education requirement"],
    "Certified": ["certified", "yes"],
    "Drug Screen": ["drug screen", "drug screem"],
    "Background Check": ["background check"],
    "Fingerprint Screening": ["fp screening", "scheduled"],
    "MDPS Training": ["mdps training"],
    "Badge Photo": ["badge photo"],
    "Notes": ["notes", "notes all calls email must be logged in ksn bh"],
}

REQUIREMENTS = [
    "Pre-hire Training", "Safety & Security", "Onboarding Introduction Email", "I-9",
    "Education Requirement", "Drug Screen", "Background Check",
    "Fingerprint Screening", "MDPS Training", "Badge Photo",
]

ARCHIVE_SHEETS = {
    "landing page", "drop downs", "template", "call back", "transitions",
    "inactivated during hiring", "newly activated", "rejected",
}


def clean_header(value):
    text = normalize_text(value)
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def normalize_text(value):
    if value is None or pd.isna(value):
        return ""
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", text).strip().lower()


def normalize_name(value):
    return re.sub(r"[^a-z0-9 ]", "", normalize_text(value))


def normalize_email(value):
    return re.sub(r"\s+", "", normalize_text(value))


def find_header_row(rows):
    for index, row in enumerate(rows[:12]):
        headers = {clean_header(value) for value in row if value is not None}
        if "applicant name" in headers and "talent email" in headers:
            return index
    return None


def canonical_column(header):
    candidate = clean_header(header)
    for canonical, aliases in FIELD_ALIASES.items():
        if any(candidate == alias or candidate.startswith(alias + " ") for alias in aliases):
            return canonical
    return None


@st.cache_data(show_spinner=False)
def read_tracker(file_bytes):
    workbook = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    records = []
    for worksheet in workbook.worksheets:
        if normalize_text(worksheet.title) in ARCHIVE_SHEETS:
            continue
        # The tracker sometimes reports thousands of formatted empty columns.
        rows = [list(row) for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row,
                                                         max_col=min(45, worksheet.max_column),
                                                         values_only=True)]
        header_index = find_header_row(rows)
        if header_index is None:
            continue
        column_map = {}
        for index, header in enumerate(rows[header_index]):
            canonical = canonical_column(header)
            if canonical and canonical not in column_map:
                column_map[canonical] = index
        if "Applicant Name" not in column_map or "Talent Email" not in column_map:
            continue
        for row in rows[header_index + 1:]:
            name = row[column_map["Applicant Name"]] if column_map["Applicant Name"] < len(row) else None
            email = row[column_map["Talent Email"]] if column_map["Talent Email"] < len(row) else None
            if not normalize_name(name) and not normalize_email(email):
                continue
            record = {field: "" for field in FIELD_ALIASES}
            for field, index in column_map.items():
                record[field] = row[index] if index < len(row) and row[index] is not None else ""
            record["Tracker Sheet"] = worksheet.title
            record["_name"] = normalize_name(name)
            record["_email"] = normalize_email(email)
            records.append(record)
    return pd.DataFrame(records)


@st.cache_data(show_spinner=False)
def read_attendance(file_bytes):
    data = pd.read_excel(io.BytesIO(file_bytes), dtype=str).fillna("")
    columns = {clean_header(column): column for column in data.columns}
    name_col = columns.get("name")
    email_col = columns.get("email")
    number_col = columns.get("n")
    if not name_col or not email_col:
        raise ValueError("The attendance file must contain NAME and EMAIL columns.")
    result = pd.DataFrame({
        "N": data[number_col] if number_col else range(1, len(data) + 1),
        "NAME": data[name_col],
        "EMAIL": data[email_col],
    })
    result["_name"] = result["NAME"].map(normalize_name)
    result["_email"] = result["EMAIL"].map(normalize_email)
    return result


def is_complete(field, value):
    value = normalize_text(value)
    if not value:
        return False
    incomplete_terms = (
        "need", "pending", "initiated", "requested", "scheduled", "will send",
        "info requested", "inconclusive", "positive", "fail", "retake", "void",
        "not cleared", "no results", "awaiting", "mailed",
    )
    if any(term in value for term in incomplete_terms):
        return False
    if field == "Education Requirement":
        return value not in {"n/a", "na", "none"}
    accepted = {
        "complete", "completed", "approved", "pass", "printed", "taken",
        "received", "recieved", "agent received", "yes", "n/a", "na",
    }
    return value in accepted


def best_match(attendee, tracker, threshold):
    email_matches = tracker[tracker["_email"].eq(attendee["_email"]) & tracker["_email"].ne("")]
    if not email_matches.empty:
        return email_matches.iloc[0], "Email match", 100
    name_matches = tracker[tracker["_name"].eq(attendee["_name"]) & tracker["_name"].ne("")]
    if not name_matches.empty:
        return name_matches.iloc[0], "Name match", 100
    if not attendee["_name"] or tracker.empty:
        return None, "Not found", 0
    scores = tracker["_name"].map(lambda name: round(SequenceMatcher(None, attendee["_name"], name).ratio() * 100))
    best_index = scores.idxmax()
    score = int(scores.loc[best_index])
    if score >= threshold:
        return tracker.loc[best_index], "Possible match – review", score
    return None, "Not found", score


def compare(attendance, tracker, threshold):
    output = []
    for _, attendee in attendance.iterrows():
        match, match_type, score = best_match(attendee, tracker, threshold)
        row = {"N": attendee["N"], "NHO Name": attendee["NAME"], "NHO Email": attendee["EMAIL"],
               "Match Result": match_type, "Match Score": score}
        if match is None:
            row.update({"Tracker Name": "", "Tracker Email": "", "Tracker Sheet": "",
                        "Current Onboarding Stage": "", "Missing Requirements": "Not found in tracker",
                        "Notes": ""})
        else:
            missing = [field for field in REQUIREMENTS if not is_complete(field, match.get(field, ""))]
            row.update({"Tracker Name": match.get("Applicant Name", ""),
                        "Tracker Email": match.get("Talent Email", ""),
                        "Tracker Sheet": match.get("Tracker Sheet", ""),
                        "Current Onboarding Stage": match.get("Current Onboarding Stage", ""),
                        "Missing Requirements": ", ".join(missing) if missing else "None",
                        "Notes": match.get("Notes", "")})
            for field in REQUIREMENTS:
                row[field] = match.get(field, "")
        output.append(row)
    return pd.DataFrame(output)


def make_excel(results):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        results.to_excel(writer, sheet_name="NHO Check", index=False)
        sheet = writer.book["NHO Check"]
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = cell.font.copy(bold=True, color="FFFFFF")
            cell.fill = cell.fill.copy(fill_type="solid", fgColor="4C9F38")
        for column in sheet.columns:
            width = min(50, max(12, max(len(str(cell.value or "")) for cell in column) + 2))
            sheet.column_dimensions[column[0].column_letter].width = width
    return buffer.getvalue()


st.title("NHO Attendance Checker")
st.caption("Compare an NHO attendance list with the Miami-Dade onboarding tracker.")

left, right = st.columns(2)
with left:
    attendance_file = st.file_uploader("1. Upload NHO attendance Excel", type=["xlsx"], key="attendance")
with right:
    tracker_file = st.file_uploader("2. Upload Miami-Dade Tracker Excel", type=["xlsx"], key="tracker")

threshold = st.slider("Minimum score for a possible name match", 70, 100, 88,
                      help="Possible matches are never treated as exact; review them before using the result.")

if attendance_file and tracker_file:
    try:
        with st.spinner("Reading and comparing files…"):
            attendance = read_attendance(attendance_file.getvalue())
            tracker = read_tracker(tracker_file.getvalue())
            results = compare(attendance, tracker, threshold)
        found = results["Match Result"].ne("Not found").sum()
        exact = results["Match Result"].isin(["Email match", "Name match"]).sum()
        review = results["Match Result"].eq("Possible match – review").sum()
        missing = results["Match Result"].eq("Not found").sum()
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Attendees", len(results)); c2.metric("Exact matches", int(exact))
        c3.metric("Review", int(review)); c4.metric("Not found", int(missing))
        st.success(f"Comparison complete: {found} of {len(results)} attendees matched the tracker.")
        filter_choice = st.segmented_control("Show", ["All", "Exact", "Review", "Not found"], default="All")
        shown = results
        if filter_choice == "Exact": shown = results[results["Match Result"].isin(["Email match", "Name match"])]
        elif filter_choice == "Review": shown = results[results["Match Result"].eq("Possible match – review")]
        elif filter_choice == "Not found": shown = results[results["Match Result"].eq("Not found")]
        st.dataframe(shown, use_container_width=True, hide_index=True,
                     column_config={"Match Score": st.column_config.ProgressColumn(min_value=0, max_value=100)})
        st.download_button("Download comparison as Excel", make_excel(results),
                           file_name="NHO_comparison.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary")
    except Exception as error:
        st.error(f"Could not compare the files: {error}")
else:
    st.info("Upload both Excel files to begin. Files are processed in memory and are not saved by the app.")
