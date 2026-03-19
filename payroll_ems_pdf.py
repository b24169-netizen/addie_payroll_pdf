import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re
import pdfplumber

st.set_page_config(page_title="EMS vs Payroll Validator", layout="wide")

st.title("EMS vs Payroll Hours Validation Tool")

ems_file = st.file_uploader("Upload EMS Monitoring Sheet", type=["xlsx"])
payroll_file = st.file_uploader("Upload Payroll PDF", type=["pdf"])


# -------- NAME CLEANING FUNCTION --------
def clean_name(name):
    name = str(name).lower()
    name = re.sub(r"[^\w\s]", "", name)
    name = name.replace(",", "")

    parts = name.split()
    if len(parts) >= 2:
        return parts[0] + parts[1]

    return name


if ems_file and payroll_file:

    # ================= EMS FILE =================
    ems_df = pd.read_excel(ems_file, header=[2, 3])

    planned = ems_df[("Planned", "Duration")]
    actual = ems_df[("Actual", "Duration")]
    employee = ems_df[("Actual", "Employee")]

    calc_df = pd.DataFrame({
        "Employee": employee,
        "Planned": planned,
        "Actual": actual
    })

    calc_df["Chosen Hours"] = calc_df[["Planned", "Actual"]].min(axis=1)

    ems_hours = (
        calc_df.groupby("Employee")["Chosen Hours"]
        .sum()
        .reset_index()
    )

    ems_hours.rename(columns={"Chosen Hours": "EMS Hours"}, inplace=True)
    ems_hours["key"] = ems_hours["Employee"].apply(clean_name)

    st.subheader("EMS Calculated Hours")
    st.dataframe(ems_hours)

    # ================= PDF PAYROLL =================
    employee_hours_map = {}
    current_employee = None
    capture_service = False
    skip_header = False

    with pdfplumber.open(payroll_file) as pdf:

        for page in pdf.pages:

            text = page.extract_text()

            if not text:
                continue

            lines = text.split("\n")

            for i, line in enumerate(lines):

                line_lower = line.lower()

                # -------- NEW EMPLOYEE --------
                if "employee address" in line_lower:

                    capture_service = False

                    if i + 1 < len(lines):
                        name_line = lines[i + 1].strip()

                        parts = name_line.split()
                        titles = ["mr", "mrs", "ms", "miss"]

                        if parts and parts[0].lower().replace(".", "") in titles:
                            parts = parts[1:]

                        if len(parts) >= 2:
                            first = parts[0]
                            last = parts[1]
                            current_employee = f"{last}, {first}"

                            if current_employee not in employee_hours_map:
                                employee_hours_map[current_employee] = 0

                    continue

                # -------- START SERVICE DETAIL --------
                if "service detail" in line_lower:
                    capture_service = True
                    skip_header = True
                    continue

                # -------- STOP CONDITIONS --------
                if "cancellation" in line_lower:
                    capture_service = False
                    continue

                # -------- SKIP HEADER ROW --------
                if skip_header:
                    skip_header = False
                    continue

                # -------- CAPTURE HOURS (FIXED LOGIC) --------
                if capture_service and current_employee:

                    # Extract decimal numbers only (ignore dates/times)
                    numbers = re.findall(r"\d+\.\d+", line)

                    if numbers:
                        try:
                            # FIRST decimal = Hours column
                            hours_value = float(numbers[0])
                            employee_hours_map[current_employee] += hours_value
                        except:
                            pass

    # -------- CONVERT TO DATAFRAME --------
    payroll_df = pd.DataFrame([
        {"Employee": emp, "Payroll Hours": hrs}
        for emp, hrs in employee_hours_map.items()
    ])

    payroll_df["key"] = payroll_df["Employee"].apply(clean_name)

    st.subheader("Detected Payroll Hours (PDF)")
    st.dataframe(payroll_df)

    # ================= MERGE =================
    if payroll_df.empty:
        st.error("No payroll records detected")
        result = ems_hours.copy()
        result["Payroll Hours"] = np.nan
    else:
        result = pd.merge(
            ems_hours,
            payroll_df,
            on="key",
            how="left",
            suffixes=("_EMS", "_Payroll")
        )

    result["Difference"] = result["EMS Hours"] - result["Payroll Hours"]

    result["Match"] = np.where(
        abs(result["Difference"]) < 0.01,
        "MATCH",
        "MISMATCH"
    )

    result = result.rename(columns={
        "Employee_EMS": "Employee"
    })

    result = result[[
        "Employee",
        "EMS Hours",
        "Payroll Hours",
        "Difference",
        "Match"
    ]]

    st.subheader("Validation Result")
    st.dataframe(result)

    # ================= EXPORT =================
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="Validation Report", index=False)

    output.seek(0)

    st.download_button(
        label="Download Validation Report",
        data=output,
        file_name="Payroll_Validation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )