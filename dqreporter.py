"""
Update 2.0

Changes: Previous versions of this report seperated CM and RA staff into seperated
reports for no, real reason other than traition.  This revised version of the
report does away with that but still maintains seperate error sheets in the final
.xlsx so that teams can easily view their mistakes and fix them accordinlyself.

General: The DQ reporter checks for missing values in basic demographics, HUD
UDE, and HUD Verification fields that are a part of all provider entries into
shelter, case management department, and housing grants.

HUD verification errors are listed as the number of fields present indicating
that there are missing fields still.  This should be reversed for intelligibility
sake in a future version.
"""

__author__ = "David Marienburg"
__maintainer__ = "David Marienburg"
__version__ = "2.0"

import pandas as pd
import numpy as np

from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

def create_dq_report(raw_staff_list, raw_data):
    """
    :param raw_staff_list: an askopenfilename file path for the
    Stafflist.xlsx
    :param raw_data: an askopenfilename file path for the QA - Demographics DQ &
    HUD Verifications v4.1.xlsx ART report
    """

    # A list of the columns the script will check for missing values.
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

    # Convert the raw ART report and the Staff File spreasheet into seperate
    # pandas dataframes.
    staff_file = pd.read_excel(raw_staff_list, sheet_name="All")
    dq_file = pd.read_excel(raw_data, header=3, sheet_name="Report 1")

    # Merge the two dataframes so that they can be compared.
    dq_named = dq_file.merge(staff_file, on="CM", how="left")

    # Convert the Entry Date column's values into a datetime.date objects.
    dq_named["Entry Date"] = dq_named["Entry Date"].dt.date

    # Create new columns for tracking the missing values.
    dq_named["Required Fields"] = 0
    dq_named["Errors"] = 0
    dq_named["Participants with Errors"] = 0

    # Loop through each dq field (first rows then columns) looking for missing
    # data, countable dq fields, or uncountable dq fields.
    for row in dq_named.index:
        for column in countable_columns:
            # Where a field has a value that isn't a - increase the errors,
            # participants with errors, and required fields columns by one.
            if pd.notnull(dq_named.loc[row, column]) and dq_named.loc[row, column] != "-":
                dq_named.loc[row, "Errors"] += 1
                dq_named.loc[row, "Participants with Errors"] = 1
                dq_named.loc[row, "Required Fields"] += 1
            # Where a field is a dash do nothing.
            elif dq_named.loc[row, column] == "-":
                pass
            # If none of the other conditions is true increase the required fields
            # count by one.
            else:
                dq_named.loc[row, "Required Fields"] += 1

    # define a dictionary that will hold the department dataframes which will
    # then be turned into sheets in the final spreadsheet.
    final_sheets = {}
    for value in dq_named["Dept"].drop_duplicates().tolist():
        final_sheets[value] = dq_named[
            (dq_named["Dept"] == value) &
            (dq_named["Errors"] > 0)
        ].drop([
            "Errors",
            "Participants with Errors",
            "Required Fields",
            "Name"
        ], axis=1)

    # Drop rows from the dq_named df where the value of the name column is DEL
    # and the Dept value is neither blank or day then save this modivied version
    # of the df to a new df named data.
    data = dq_named[
        (dq_named["Name"] != "DEL") &
        (dq_named["Dept"] != "Day") &
        (dq_named["Dept"] != "DEL") &
        dq_named["Dept"].notnull()
    ]

    # Create a pivot table from the data dataframe to summarize the errors by
    # department and staff name.
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

    # Add the calculated columns to the staff_summary pivot table.
    staff_summary["Error Rate"] = staff_summary["Participants with Errors"]/staff_summary["CTID"]
    staff_summary["Errors Per Participants"] = staff_summary["Errors"]/staff_summary["CTID"]
    staff_summary["Error Rate Per Required Field"] = staff_summary["Errors"]/staff_summary["Required Fields"]

    # Create a pivot table from the dataframe to summarize the errors by
    # department only.
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
    # Add the calculated columns to the dept_summary pivot table.
    dept_summary["Error Rate"] = dept_summary["Participants with Errors"]/dept_summary["CTID"]
    dept_summary["Errors Per Participants"] = dept_summary["Errors"]/dept_summary["CTID"]
    dept_summary["Error Rate Per Required Field"] = dept_summary["Errors"]/dept_summary["Required Fields"]

    # Initialize the writer object.
    writer = pd.ExcelWriter(asksaveasfilename(title="Save the CM Report"), engine="xlsxwriter")

    # Add the two summary sheets to the writer object.
    dept_summary.to_excel(writer, sheet_name="Dept Summary")
    staff_summary.to_excel(writer, sheet_name="Staff Summary")

    # Loop through the keys in the final_sheets dictionary adding a new sheet to
    # the writer object for each key filled with the associated values and named
    # using the key name.
    for key in final_sheets.keys():
        final_sheets[key].to_excel(writer, sheet_name="{}".format(key), index=False)

    # Create a raw data sheet.
    dq_named.to_excel(writer, sheet_name="Raw Data", index=False)

    # Save the writer object
    writer.save()

if __name__ == "__main__":
    staff = askopenfilename(title="Open the staff names spreadsheet")
    data = askopenfilename(title="Open the Demographics and DataQuality ART Report")
    create_dq_report(staff, data)
