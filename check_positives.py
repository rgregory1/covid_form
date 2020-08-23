import pygsheets
from pprint import pprint

# import os
import yagmail
import credentials
import datetime
import maya

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


def get_school_lists():

    # open up google sheet to see if new staff have been added
    gc = pygsheets.authorize(outh_file="client_secret.json")
    school_covid_lists = gc.open_by_key("1ZIisaL_CQkbblkh7Furrd9J_MIQMeViWLrewgJndIns")
    swanton_list = school_covid_lists.worksheet_by_title("Swanton")

    # download all data from sheet as cell_matrix
    cell_matrix = swanton_list.get_all_values(returnas="matrix")
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
    swanton_checklist = []

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
            swanton_checklist.append(line_dict)

    # remove first line with headings as data
    swanton_checklist.pop(0)

    # print(swanton_checklist)
    print("Swanton list complete")

    return swanton_checklist


swanton_checklist = get_school_lists()

# # test dict so I don't send everyone emails
# recipients_dict = {
#     "Central Office": "russell.gregory@mvsdschools.org",
#     "Swanton Elementary: Central Building": "russell.gregory@mvsdschools.org",
#     "MVU": "russell.gregory@mvsdschools.org",
#     "MVU-Connect": "russell.gregory@mvsdschools.org",
#     "Franklin Central School": "russell.gregory@mvsdschools.org",
#     "Highgate Elementary": "russell.gregory@mvsdschools.org",
#     "Swanton Elementary: Babcock Building": [
#         "russell.gregory@mvsdschools.org",
#         "robert.gervais@mvsdschools.org",
#     ],
# }

# setup recipients emails by scchool
recipients_dict = {
    "Central Office": "Pierrette.Bouchard@mvsdschools.org",
    "Swanton Elementary: Central Building": [
        "Danielle.Loiselle@mvsdschools.org",
        "russell.gregory@mvsdschools.org",
    ],
    "MVU": ["Alissa.Graves@mvsdschools.org", "Lynn.Billado@mvsdschools.org"],
    "MVU-Connect": "russell.gregory@mvsdschools.org",
    "Franklin Central School": "alita.boomhower@mvsdschools.org",
    "Highgate Elementary": "Jennifer.Gagne@mvsdschools.org",
    "Swanton Elementary: Babcock Building": "Wendy.Culligan@mvsdschools.org",
}


def check_for_new_staff():
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

    # put cell_matrix list of lists into a list of dictionaries
    for count, row in enumerate(cell_matrix):
        if row[0] != "":
            if row[9] == "":
                # create dict from list of info
                line_dict = dict(zip(dict_key_list, row))

                # add count so I can add x to appropriate row later
                line_dict["row_number"] = count + 1

                # create timestap so I can tell if form filled today
                line_dict["Timestamp"] = maya.parse(line_dict["Timestamp"]).datetime()

                # create final list of entered form data
                worksheet_data.append(line_dict)

    print("Check complete")

    return worksheet_data, initial_form_sheet


def email_nurse(staff, recipients_dict):

    # get list of schools the employee works at
    schools = staff["Buildings"]

    # split string of schools into a list
    schools_list = schools.split(", ")

    # Send email to each school respondant is attending
    for school in schools_list:
        for key in recipients_dict:
            if key in school:

                nurse_address = recipients_dict[key]

                contents = f"""\nATTENTION:\n\n
                A staff member at your school, {staff['Name']}, has reported 'YES' to 
                one or more of the questions on the COVID-19 Ready to Work Screening.
                
                Please find the response below:

                Name: {staff['Name']}
                Buildings Working in Today: {staff['Buildings']}
                Form Submitted: {staff['Timestamp']}
                """
                html = f"""<table style="border-collapse: collapse;">
                        <tbody>
                            <tr style="background-color:#D4DADC;">
                                <td> Was your temperature over 100.4º F or 38º C before coming to work today? </td>
                                <td> {staff['Temp']}</td>
                            
                            </tr>
                            <tr>
                                <td> In the last 14 days, have you traveled outside your normal, daily routine 
                                to a location that has been identified as a Hot Spot or Red Zone (400 or more 
                                confirmed cases per 1 million residents)?</td>
                                <td> {staff['Travel']}</td>
                            </tr>
                            <tr style="background-color:#D4DADC;">
                                <td> Do you have a new or worsening onset of any of the following symptoms: 
                                fever, cough, shortness of breath, runny nose, sore throat, chills, body aches, 
                                fatigue, headache, loss of taste/smell, eye drainage, or congestion or any other 
                                COVID related symptoms?</td>
                                <td> {staff['Symptoms']}</td>
                            </tr>
                            <tr>
                                <td> Have you been exposed to someone being tested for COVID-19 or who has 
                                symptoms compatible with COVID-19?</td>
                                <td> {staff['Exposure']}</td>
                            </tr>
                            <tr style="background-color:#D4DADC;">
                                <td> Are any members of your household or close contact in quarantine for 
                                exposure to COVID-19?</td>
                                <td> {staff['Household']}</td>
                            </tr>
                        </tbody>
                    </table>"""
                yag.send(nurse_address, "Covid Form Alert", [contents, html])


def email_att_list(swanton_checklist, filled_names_set, filled_pins_set):

    # begin constructing list of those that filled it in and those that didn't
    for staff in swanton_checklist:

        # if staff memember is in the swanton building check against swanton_checklist
        if staff["Name"] in filled_names_set or staff["ID"] in filled_pins_set:
            staff["filled"] = True

    disclaimer = "This information is valid as of " + str(temp_timestamp) + "\n\n\n"
    filled_it_out = "The Following folks filled out the form today\n----------------\n"
    did_not_fill_it_out = (
        "\n\n\nThe Following folks did not fill out the form today\n----------------\n"
    )

    for staff in swanton_checklist:
        if staff["filled"] == True:
            filled_it_out = filled_it_out + staff["Name"].title() + "\n"

    for staff in swanton_checklist:
        if staff["filled"] == False:
            did_not_fill_it_out = did_not_fill_it_out + staff["Name"].title() + "\n"

    # # test email information
    # yag.send(
    #     "russell.gregory@mvsdschools.org",
    #     "COVID Daily Form Roll Call",
    #     [disclaimer, filled_it_out, did_not_fill_it_out],
    # )
    yag.send(
        [
            "Justina.Jennett@mvsdschools.org",
            "dawn.tessier@mvsdschools.org",
            "Mary.Ellis@mvsdschools.org",
            "russell.gregory@mvsdschools.org",
        ],
        "COVID Daily Form Roll Call",
        [disclaimer, filled_it_out, did_not_fill_it_out],
    )

    return


# check for new staff
worksheet_data, initial_form_sheet = check_for_new_staff()

# create sets of filled in info
filled_names_set = set()
filled_pins_set = set()

for staff in worksheet_data:
    # check if form was filled in today
    if staff["Timestamp"].date() == datetime.datetime.today().date():
        filled_names_set.add(staff["Name"].lower())
        if staff["ID"]:
            filled_pins_set.add(staff["ID"])

# check for new staff
if len(worksheet_data) == 0:
    # empty list
    print("No new responses found")

else:
    # list contains items
    print("New response found\n")
    is_new_response = True

    for staff in worksheet_data:
        if "YES" in staff.values():
            print(f"Sending mail for {staff['Name']}\n")
            email_nurse(staff, recipients_dict)

        # mark new staff member as processed with X in column J
        mark_as_finished_cell = "J" + str(staff["row_number"])
        initial_form_sheet.update_value(mark_as_finished_cell, "X")

    email_att_list(swanton_checklist, filled_names_set, filled_pins_set)


print("\n\nfinished")
