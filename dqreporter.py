import pandas as pd
import numpy as np

from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

def create_cm_dq_report(raw_staff, raw_data):
    # add a line dropping rows where department == Del

    countable_columns = [
            "SSID",
            "SSID Type",
            "Vet Status",
            "DoB",
            "DoB Type",
            "Race",
            "Race-Additional",
            "Ethnicity",
            "Gender",
            "Relationship to HoH",
            "Client Location",
            "Prior Residence",
            "LoS",
            "Date Homelessness Started",
            "Times Homeless",
            "Total Months Homeless",
            "Income From Any Source",
            "Income Verification",
            "Covered By Insurance",
            "Insurance Verification",
            "Non-Cash Benefits From Any Source",
            "Non-Cash Benefits Verification",
            "Does the Client Have a Disabling Conditon",
            "Disability Type",
            "DV Survivor",
            "DV Date",
            "DV Fleeing"
    ]
    provider_dict = {
        "Transition Projects (TPI) - ACCESS - CM(5471)": "Outreach",
        "Transition Projects (TPI) - Residential - CM(5473)": "Residential",
        "Transition Projects (TPI) - Retention - CM(5472)": "Retention",
        "Transition Projects (TPI) - SSVF_C15-OR-501A - Homeless Prevention (VA) - SP(4803)": "SSVF",
        "Transition Projects (TPI) - SSVF_Renewal 15-ZZ-127 - Homeless Prevention (VA) - SP(4801)": "SSVF",
        "Transition Projects (TPI) - SSVF_C15-OR-501A - Rapid Re-Housing (VA) - SP(4804)": "SSVF",
        "Transition Projects (TPI) - SSVF_Renewal 15-ZZ-127 Rapid Re-Housing (VA) - SP(4802)": "SSVF"
    }

    staff_file = pd.read_excel(raw_staff, sheet_name="CM")
    dq_file = pd.read_excel(raw_data, header=3, sheet_name="Report 1")

    dq_named = dq_file.merge(staff_file, on="CM", how="left")
    dq_named["Required Fields"] = 0
    dq_named["Errors"] = 0
    dq_named["Participants with Errors"] = 0

    for row in dq_named.index:
        for column in countable_columns:
            if pd.notnull(dq_named.loc[row, column]) and dq_named.loc[row, column] != "-":
                dq_named.loc[row, "Errors"] += 1
                dq_named.loc[row, "Participants with Errors"] = 1
                dq_named.loc[row, "Required Fields"] += 1
            elif dq_named.loc[row, column] != "-":
                dq_named.loc[row, "Required Fields"] += 1
            else:
                pass

    outreach = dq_named[(dq_named["Dept"] == "OUT") & (dq_named["Errors"] > 0)].drop([
        "Department",
        "Errors",
        "Participants with Errors",
        "Required Fields"
    ], axis=1)
    res = dq_named[(dq_named["Dept"] == "RES") & (dq_named["Errors"] > 0)].drop([
        "Department",
        "Errors",
        "Participants with Errors",
        "Required Fields"
    ], axis=1)
    ret = dq_named[(dq_named["Dept"] == "RET") & (dq_named["Errors"] > 0)].drop([
        "Department",
        "Errors",
        "Participants with Errors",
        "Required Fields"
    ], axis=1)
    ssvf = dq_named[(dq_named["Dept"] == "SSVF") & (dq_named["Errors"] > 0)].drop([
        "Department",
        "Errors",
        "Participants with Errors",
        "Required Fields"
    ], axis=1)
    data = dq_named[~(dq_named["Name"] == "DEL")]

    staff_summary = pd.pivot_table(
        data,
        index=["Dept", "Name"],
        values=["CTID", "Required Fields", "Participants with Errors", "Errors"],
        aggfunc={
            "CTID": len,
            "Errors": np.sum,
            "Participants with Errors": np.sum,
            "Required Fields": np.sum
        }
    )
    staff_summary["Error Rate"] = staff_summary["Participants with Errors"]/staff_summary["CTID"]
    staff_summary["Errors Per Participants"] = staff_summary["Errors"]/staff_summary["CTID"]
    staff_summary["Error Rate Per Required Field"] = staff_summary["Errors"]/staff_summary["Required Fields"]

    dept_summary = pd.pivot_table(
        data,
        index=["Dept"],
        values=["CTID", "Required Fields", "Participants with Errors", "Errors"],
        aggfunc={
            "CTID": len,
            "Errors": np.sum,
            "Participants with Errors": np.sum,
            "Required Fields": np.sum
        }
    )
    dept_summary["Error Rate"] = dept_summary["Participants with Errors"]/dept_summary["CTID"]
    dept_summary["Errors Per Participants"] = dept_summary["Errors"]/dept_summary["CTID"]
    dept_summary["Error Rate Per Required Field"] = dept_summary["Errors"]/dept_summary["Required Fields"]

    writer = pd.ExcelWriter(asksaveasfilename(title="Save the CM Report"), engine="xlsxwriter")
    dept_summary.to_excel(writer, sheet_name="Dept Summary")
    staff_summary.to_excel(writer, sheet_name="Staff Summary")
    outreach.to_excel(writer, sheet_name="OUT", index=False)
    res.to_excel(writer, sheet_name="RES", index=False)
    ret.to_excel(writer, sheet_name="RET", index=False)
    ssvf.to_excel(writer, sheet_name="SSVF", index=False)
    data.to_excel(writer, sheet_name="Data Processed", index=False)
    writer.save()

def create_ra_dq_report(raw_staff, raw_data):
    # add a line dropping rows where department == Del

    countable_columns = [
            "SSID",
            "SSID Type",
            "Vet Status",
            "DoB",
            "DoB Type",
            "Race",
            "Race-Additional",
            "Ethnicity",
            "Gender",
            "Relationship to HoH",
            "Client Location",
            "Prior Residence",
            "LoS",
            "Date Homelessness Started",
            "Times Homeless",
            "Total Months Homeless",
            "Income From Any Source",
            "Income Verification",
            "Covered By Insurance",
            "Insurance Verification",
            "Non-Cash Benefits From Any Source",
            "Non-Cash Benefits Verification",
            "Does the Client Have a Disabling Conditon",
            "Disability Type",
            "DV Survivor",
            "DV Date",
            "DV Fleeing"
    ]
    provider_dict = {
        "Transition Projects (TPI) - Clark Center - SP(25)": "CC",
        "ZZ - Transition Projects (TPI) - Columbia Shelter (Do not use after 4/25/18) (Level 6)(5857) ": "COL",
        "Transition Projects (TPI) - Columbia Shelter(6527)": "COL",
        "Transition Projects (TPI) - Doreen's Place - SP(28)": "DP",
        "Transition Projects (TPI) - Hansen Emergency Shelter - SP(5588)": "H",
        "Transition Projects (TPI) - Jean's Place L1 - SP(29)": "JP",
        "Transition Projects (TPI) - SOS Shelter(2712)": "SOS",
        "Transition Projects (TPI) - VA Grant Per Diem (inc. Doreen's Place GPD) - SP(3189)": "DP",
        "Transition Projects (TPI) - Willamette Center(5764)": "W",
        "Transition Projects (TPI) - 5th Avenue Shelter(6281)": "5th"
    }

    staff_file = pd.read_excel(raw_staff, sheet_name="RA")
    dq_file = pd.read_excel(
        raw_data,
        header=3,
        sheet_name="Report 1"
    )

    dq_file["Shelter"] = dq_file["Department"].apply(lambda x: provider_dict.get(x))
    dq_file["Required Fields"] = 0
    dq_file["Errors"] = 0
    dq_file["Participants with Errors"] = 0

    for row in dq_file.index:
        for column in countable_columns:
            if pd.notnull(dq_file.loc[row, column]) and dq_file.loc[row, column] != "-":
                dq_file.loc[row, "Errors"] += 1
                dq_file.loc[row, "Participants with Errors"] = 1
                dq_file.loc[row, "Required Fields"] += 1
            elif dq_file.loc[row, column] != "-":
                dq_file.loc[row, "Required Fields"] += 1
            else:
                pass

    dq_named = dq_file.merge(staff_file, on="CM", how="left")

    cc = dq_named[(dq_file["Shelter"] == "CC") & (dq_named["Errors"] > 0)].drop([
        "Department",
        "Errors",
        "Participants with Errors",
        "Required Fields"
    ], axis=1)
    col = dq_named[(dq_file["Shelter"] == "COL") & (dq_named["Errors"] > 0)].drop([
        "Department",
        "Errors",
        "Participants with Errors",
        "Required Fields"
    ], axis=1)
    dp = dq_named[(dq_file["Shelter"] == "DP") & (dq_named["Errors"] > 0)].drop([
        "Department",
        "Errors",
        "Participants with Errors",
        "Required Fields"
    ], axis=1)
    jp = dq_named[(dq_file["Shelter"] == "JP") & (dq_named["Errors"] > 0)].drop([
        "Department",
        "Errors",
        "Participants with Errors",
        "Required Fields"
    ], axis=1)
    w = dq_named[(dq_file["Shelter"] == "W") & (dq_named["Errors"] > 0)].drop([
        "Department",
        "Errors",
        "Participants with Errors",
        "Required Fields"
    ], axis=1)
    sos = dq_named[(dq_file["Shelter"] == "SOS") & (dq_named["Errors"] > 0)].drop([
        "Department",
        "Errors",
        "Participants with Errors",
        "Required Fields"
    ], axis=1)
    h = dq_named[(dq_file["Shelter"] == "H") & (dq_named["Errors"] > 0)].drop([
        "Department",
        "Errors",
        "Participants with Errors",
        "Required Fields"
    ], axis=1)
    fifth = dq_named[(dq_file["Shelter"] == "5th") & (dq_named["Errors"] > 0)].drop([
        "Department",
        "Errors",
        "Participants with Errors",
        "Required Fields"
    ], axis=1)

    staff_summary = pd.pivot_table(
        dq_named,
        index=["Shelter", "CM"],
        values=["CTID", "Required Fields", "Participants with Errors", "Errors"],
        aggfunc={
            "CTID": len,
            "Errors": np.sum,
            "Participants with Errors": np.sum,
            "Required Fields": np.sum
        }
    )
    staff_summary["Error Rate"] = staff_summary["Participants with Errors"]/staff_summary["CTID"]
    staff_summary["Errors Per Participants"] = staff_summary["Errors"]/staff_summary["CTID"]
    staff_summary["Error Rate Per Required Field"] = staff_summary["Errors"]/staff_summary["Required Fields"]

    dept_summary = pd.pivot_table(
        dq_named,
        index=["Shelter"],
        values=["CTID", "Required Fields", "Participants with Errors", "Errors"],
        aggfunc={
            "CTID": len,
            "Errors": np.sum,
            "Participants with Errors": np.sum,
            "Required Fields": np.sum
        }
    )
    dept_summary["Error Rate"] = dept_summary["Participants with Errors"]/dept_summary["CTID"]
    dept_summary["Errors Per Participants"] = dept_summary["Errors"]/dept_summary["CTID"]
    dept_summary["Error Rate Per Required Field"] = dept_summary["Errors"]/dept_summary["Required Fields"]

    writer = pd.ExcelWriter(asksaveasfilename(title="Save the Shelter Report"), engine="xlsxwriter")
    dept_summary.to_excel(writer, sheet_name="Dept Summary")
    staff_summary.to_excel(writer, sheet_name="Staff Summary")
    fifth.to_excel(writer, sheet_name="5th", index=False)
    cc.to_excel(writer, sheet_name="CC", index=False)
    col.to_excel(writer, sheet_name="COL", index=False)
    dp.to_excel(writer, sheet_name="DP", index=False)
    jp.to_excel(writer, sheet_name="JP", index=False)
    w.to_excel(writer, sheet_name="W", index=False)
    sos.to_excel(writer, sheet_name="SOS", index=False)
    h.to_excel(writer, sheet_name="H", index=False)
    dq_named.to_excel(writer, sheet_name="Data Processed", index=False)
    writer.save()


if __name__ == "__main__":
    staff = askopenfilename(title="Open the staff names spreadsheet")
    data = askopenfilename(title="Open the Demographics and DataQuality ART Report")
    create_ra_dq_report(staff, data)
    create_cm_dq_report(staff, data)
