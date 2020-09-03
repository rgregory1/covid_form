import pygsheets
from pprint import pprint

# import os
import yagmail
import credentials
import datetime

# import maya

# setup gmail link
gmail_user = credentials.gmail_user
gmail_password = credentials.gmail_password
yag = yagmail.SMTP(gmail_user, gmail_password)


# get timestamp for log
temp_timestamp = datetime.datetime.now()
print(str(temp_timestamp))
print("\n")
print("checking new daily covid form entries")
print("\n")


def get_school_lists(school_name):

    # open up google sheet to see if new staff have been added
    gc = pygsheets.authorize(outh_file="client_secret.json")
    school_covid_lists = gc.open_by_key("1ZIisaL_CQkbblkh7Furrd9J_MIQMeViWLrewgJndIns")
    school_list = school_covid_lists.worksheet_by_title(school_name)

    # download all data from sheet as cell_matrix
    cell_matrix = school_list.get_all_values(returnas="matrix")
    # print(cell_matrix)

    # gather 'keys' for new dict from 1st row in sheet
    dict_key_list = [x for x in cell_matrix[0]]
    # pprint(dict_key_list)
    dict_key_list = [
        "First Name",
        "Last Name",
        "Work Email",
        "ID",
    ]

    # initialize dict for data
    school_checklist = []

    # put cell_matrix list of lists into a list of dictionaries
    for row in cell_matrix:
        if row[0] != "":
            # create dict from list of lists
            line_dict = dict(zip(dict_key_list, row))
            # concatinate names into one lower string for comparison
            first = line_dict["First Name"].lower()
            last = line_dict["Last Name"].lower()
            line_dict["Name"] = first + " " + last

            # remove unneeded namee pairs
            line_dict.pop("First Name")
            line_dict.pop("Last Name")

            # add boolean for list creation later
            line_dict["filled"] = False

            # build final list for checking against
            school_checklist.append(line_dict)

    # remove first line with headings as data
    school_checklist.pop(0)

    # print(swanton_checklist)
    print(f"{school_name} list complete")

    return school_checklist


highgate_checklist = get_school_lists("HES")
swanton_checklist = get_school_lists("Swanton")
central_checklist = get_school_lists("Central")
franklin_checklist = get_school_lists("FCS")


# # setup recipients emails by scchool for attendance
# att_emails_dict = {
#     "Central": "russell.gregory@mvsdschools.org",
#     "Swanton": "russell.gregory@mvsdschools.org",
#     "Franklin": "russell.gregory@mvsdschools.org",
#     "Highgate": "russell.gregory@mvsdschools.org",
# }


# setup recipients emails by scchool for attendance
att_emails_dict = {
    "Central": "Pierrette.Bouchard@mvsdschools.org",
    "Swanton": [
        "Justina.Jennett@mvsdschools.org",
        "dawn.tessier@mvsdschools.org",
        "Mary.Ellis@mvsdschools.org",
        "russell.gregory@mvsdschools.org",
    ],
    "Franklin": "kathy.ovitt@mvsdschools.org",
    "Highgate": ["amber.LaFar@mvsdschools.org", "russell.gregory@mvsdschools.org",],
}


def check_for_roll_call():
    print("Starting check of daily covid form...")

    # open up google sheet to see if new staff have been added
    gc = pygsheets.authorize(outh_file="client_secret.json")
    initial_form_wb = gc.open_by_key("1Rn7BLeVfg9UyAlAufHiogonyQhogrQS8zkH03Ne9Ark")
    initial_form_sheet = initial_form_wb.worksheet_by_title("Daily_Covid_Responses")

    # download all data from sheet as cell_matrix
    cell_matrix = initial_form_sheet.get_all_values(returnas="matrix")
    # print(cell_matrix)

    # gather 'keys' for new dict from 1st row in sheet
    dict_key_list = [x for x in cell_matrix[0]]
    # pprint(dict_key_list)
    dict_key_list = [
        "Timestamp",
        "Name",
        "Buildings",
        "Temp",
        "Travel",
        "Symptoms",
        "Exposure",
        "Household",
        "ID",
    ]

    # initialize dict for data
    worksheet_data = []

    # create sets of filled in info
    filled_names_set = set()
    filled_pins_set = set()

    # removed first line of cell matrix containing headers
    cell_matrix.pop(0)

    # put cell_matrix list of lists into a list of dictionaries
    for count, row in enumerate(cell_matrix):
        if row[0] != "":
            # create dict from list of info
            line_dict = dict(zip(dict_key_list, row))
            # strip out all extra spaces from name field
            line_dict["Name"] = " ".join(line_dict["Name"].split())
            # create timestap so I can tell if form filled today
            # line_dict["Timestamp"] = maya.parse(line_dict["Timestamp"]).datetime()
            line_dict["Timestamp"] = datetime.datetime.strptime(
                line_dict["Timestamp"], "%m/%d/%Y %H:%M:%S"
            )

            # create final list of entered form data
            worksheet_data.append(line_dict)

    for staff in worksheet_data:
        # check if form was filled in today
        if staff["Timestamp"].date() == datetime.datetime.today().date():
            filled_names_set.add(staff["Name"].lower())
            if staff["ID"]:
                filled_pins_set.add(staff["ID"])

    # sorted list of the set for debugging purposes
    alpha_filled_names = sorted(filled_names_set)

    # # for test purposses only
    # print(alpha_filled_names)

    print("Check complete")

    return filled_names_set, filled_pins_set, alpha_filled_names


def email_att_list(
    school_checklist, email_list, school, filled_names_set, filled_pins_set
):

    print(f"beginning {school} attendance check")
    # begin constructing list of those that filled it in and those that didn't
    for staff in school_checklist:

        # if staff memember is in the swanton building check against swanton_checklist
        if staff["Name"] in filled_names_set or staff["ID"] in filled_pins_set:
            staff["filled"] = True

    disclaimer = "This information is valid as of " + str(temp_timestamp) + "\n\n\n"
    filled_it_out = "The Following folks filled out the form today\n----------------\n"
    did_not_fill_it_out = (
        "\n\n\nThe Following folks did not fill out the form today\n----------------\n"
    )

    for staff in school_checklist:
        if staff["filled"] == True:
            filled_it_out = filled_it_out + staff["Name"].title() + "\n"

    for staff in school_checklist:
        if staff["filled"] == False:
            did_not_fill_it_out = did_not_fill_it_out + staff["Name"].title() + "\n"

    # # test email information
    # yag.send(
    #     "russell.gregory@mvsdschools.org",
    #     "COVID Daily Form Roll Call",
    #     [disclaimer, filled_it_out, did_not_fill_it_out],
    # )
    yag.send(
        email_list,
        "COVID Daily Form Roll Call",
        [disclaimer, filled_it_out, did_not_fill_it_out],
    )

    print(f"email sent for {school}")
    return


# check for new staff
filled_names_set, filled_pins_set, alpha_filled_names = check_for_roll_call()


# check for new staff
if len(filled_names_set) == 0:
    # empty list
    print("No new responses found")

else:
    # list contains items
    print("New response found\n")
    is_new_response = True

    email_att_list(
        swanton_checklist,
        att_emails_dict["Swanton"],
        "Swanton",
        filled_names_set,
        filled_pins_set,
    )

    email_att_list(
        highgate_checklist,
        att_emails_dict["Highgate"],
        "Highgate",
        filled_names_set,
        filled_pins_set,
    )

    email_att_list(
        franklin_checklist,
        att_emails_dict["Franklin"],
        "Franklin",
        filled_names_set,
        filled_pins_set,
    )

    email_att_list(
        central_checklist,
        att_emails_dict["Central"],
        "Central",
        filled_names_set,
        filled_pins_set,
    )


print("\n\nfinished")
