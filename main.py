# ------SMIDST (Social Media Intern's Data Searching Tool)
# ------Created by Himanshu Kawale (9175468474)

from __future__ import print_function
from tkinter.constants import LEFT
from googleapiclient.discovery import build
from tkinter import Message, Label, Frame, Button, Tk, Entry
from google.oauth2 import service_account
from webbrowser import open
import time
from functools import lru_cache

url = ""
no_internet = 0
HR_contact = None


def set_HR_url():
    global url
    url = f"http://wa.me/91{HR_contact}"
    openweb()


def openweb():
    new = 1
    open(url, new=new)


try:
    # If modifying these scopes, delete the file keys.json.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    SERVICE_ACCOUNT_FILE = 'keyfile.json'
    RANGE_SHEET_RANGE = ["Sheet data"]
    creds = None
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    # copy from sheet url
    SAMPLE_SPREADSHEET_ID = '1Q7OZFo1zm32Tm_Ekj3FKcCSB9qY4rRIBnltbrtKJ17w'
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
except:
    # ("error in key")
    pass


def HR_search():
    global sheet
    global details
    global RANGE_SHEET_RANGE
    global SAMPLE_SPREADSHEET_ID
    RANGE_NAME = []
    HR_contact, BACKGROUND_COLOR = "", "white"

    result = sheet.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                       ranges=RANGE_SHEET_RANGE, includeGridData=True).execute()
    row = 0

    while 1:
        try:
            for key, value in result["sheets"][0]["data"][0]["rowData"][row]["values"][1]["userEnteredValue"].items():
                if value != "":
                    RANGE_NAME.append(str(value))
        except:
            break
        row = row + 1

    sheet_range, sheet_num, data_avilable = 0, 0, 0

    result = sheet.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                       ranges=RANGE_NAME, includeGridData=True).execute()

    sheet_range = len(result["sheets"])

    for sheet_num in range(sheet_range):
        row, col, end_row_count, end_of_sheet, HR_name = 1, 0, 0, 0, None

        while 1:
            try:
                HR_name = str(result["sheets"][sheet_num]["data"][0]["rowData"]
                              [row]["values"][0]["userEnteredValue"]['stringValue']).lower()
                
                if HR_name in str(details[3]).lower():
                    BACKGROUND_COLOR = result["sheets"][sheet_num]["data"][0]["rowData"][
                        row]["values"][col]["effectiveFormat"]["backgroundColor"]

                    if str(BACKGROUND_COLOR) == "{'red': 1, 'blue': 1}":
                        row = row + 1
                        break

                    HR_contact = result["sheets"][sheet_num]["data"][0]["rowData"][row]["values"][2]["userEnteredValue"]["numberValue"]
                    data_avilable = 1
                    break

                else:
                    row = row + 1
            except:
                row = row + 1
                end_row_count = end_row_count + 1
                if end_row_count == 10:
                    end_of_sheet = 1
                if end_of_sheet:
                    break
        if data_avilable:
            break
    return HR_contact, BACKGROUND_COLOR, data_avilable, HR_name


@lru_cache(maxsize=5)
def range_creator():
    global sheet
    global no_internet
    global RANGE_SHEET_RANGE
    global SAMPLE_SPREADSHEET_ID
    RANGE_NAME = []
    try:
        result = sheet.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                           ranges=RANGE_SHEET_RANGE, includeGridData=True).execute()
    except:
        no_internet = 1
        return 0
    row = 0

    while 1:
        try:
            for key, value in result["sheets"][0]["data"][0]["rowData"][row]["values"][0]["userEnteredValue"].items():
                if value != "":
                    RANGE_NAME.append(str(value))
        except:
            break
        row = row + 1
    return RANGE_NAME


@lru_cache(maxsize=5)
def result_creator():
    RANGE_NAME = range_creator()
    if RANGE_NAME == 0:  # for No internet connection
        return 0
    result = sheet.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                       ranges=RANGE_NAME, includeGridData=True).execute()
    return result


@lru_cache(maxsize=5)
def main(user_number):
    global sheet

    result = result_creator()
    if result == 0:
        return 0

    sheet_range, sheet_num, data_avilable = 0, 0, 0
   

    sheet_range = len(result["sheets"])

    INTERN_NUM_DICT, INTERN_EXTRA_DICT, INTERN_NAME_DICT, BACKGROUND_COLOR, INTERN_EMAIL_DICT = "", "", "", "", ""

    for sheet_num in range(sheet_range):
        row, col, end_row_count, end_of_sheet, end_of_col = 0, 0, 0, 0, 0

        while 1:
            try:
                if user_number in str(result["sheets"][sheet_num]["data"][0]["rowData"][row]["values"][1]["userEnteredValue"].items()):
                    BACKGROUND_COLOR = result["sheets"][sheet_num]["data"][0]["rowData"][
                        row]["values"][col]["effectiveFormat"]["backgroundColor"]

                    end_col_count, end_of_col, value = 0, 0, 1
                    while 1:
                        if str(BACKGROUND_COLOR) == "{'red': 1, 'blue': 1}":
                            row = row + 1
                            break
                        try:
                            while 1:
                                VALUE_DICT = result["sheets"][sheet_num]["data"][0]["rowData"][row]["values"][col]["userEnteredValue"]
                                if value == 1:
                                    INTERN_NAME_DICT = VALUE_DICT
                                    value = value + 1
                                elif value == 2:
                                    INTERN_NUM_DICT = VALUE_DICT
                                    value = value + 1
                                elif value > 2:
                                    if VALUE_DICT != "":
                                        if "@" in str(VALUE_DICT):
                                            INTERN_EMAIL_DICT = VALUE_DICT
                                        else:
                                            if "http" not in str(VALUE_DICT):
                                                if "CONCATENATE" not in str(VALUE_DICT):
                                                    if "N/A" not in str(VALUE_DICT):
                                                        if len(str(VALUE_DICT)) < 100:
                                                            for key, val in VALUE_DICT.items():
                                                                if "yes" not in val.lower():
                                                                    if "/" not in val:
                                                                        INTERN_EXTRA_DICT = "\n".join(
                                                                            [INTERN_EXTRA_DICT, val])
                                    value = value + 1
                                col = col + 1
                        except:
                            col = col + 1
                            end_col_count = end_col_count + 1
                            if end_col_count == 15:
                                end_of_col = 1
                                # end_col_count = 0
                            if end_of_col:
                                # end_of_col = 0
                                break
                    if str(BACKGROUND_COLOR) != "{'red': 1, 'blue': 1}":
                        data_avilable = 1
                        break
                else:
                    row = row + 1
            except:
                row = row + 1
                end_row_count = end_row_count + 1
                if end_row_count == 10:
                    end_of_sheet = 1
                if end_of_sheet:
                    break
        if data_avilable:
            break

    return INTERN_NAME_DICT, INTERN_NUM_DICT, INTERN_EMAIL_DICT, INTERN_EXTRA_DICT, BACKGROUND_COLOR, data_avilable


@lru_cache(maxsize=5)
def main_byName(user_name):
    global sheet
    result = result_creator()

    if result == 0:
        return 0

    sheet_range, sheet_num, data_avilable = 0, 0, 0

    sheet_range = len(result["sheets"])

    INTERN_NUM_DICT, INTERN_EXTRA_DICT, INTERN_NAME_DICT, BACKGROUND_COLOR, INTERN_EMAIL_DICT = "", "", "", "", ""

    for sheet_num in range(sheet_range):
        row, col, end_row_count, end_of_sheet, end_of_col = 0, 0, 0, 0, 0

        while 1:
            try:
                if user_name.lower() in str(result["sheets"][sheet_num]["data"][0]["rowData"][row]["values"][0]["userEnteredValue"].items()).lower():
                    BACKGROUND_COLOR = result["sheets"][sheet_num]["data"][0]["rowData"][
                        row]["values"][col]["effectiveFormat"]["backgroundColor"]

                    end_col_count, end_of_col, value = 0, 0, 1
                    while 1:
                        if str(BACKGROUND_COLOR) == "{'red': 1, 'blue': 1}":
                            row = row + 1
                            break
                        try:
                            while 1:
                                VALUE_DICT = result["sheets"][sheet_num]["data"][0]["rowData"][row]["values"][col]["userEnteredValue"]
                                if value == 1:
                                    INTERN_NAME_DICT = VALUE_DICT
                                    value = value + 1
                                elif value == 2:
                                    INTERN_NUM_DICT = VALUE_DICT
                                    value = value + 1
                                elif value > 2:
                                    if VALUE_DICT != "":
                                        if "@" in str(VALUE_DICT):
                                            INTERN_EMAIL_DICT = VALUE_DICT
                                        else:
                                            if "http" not in str(VALUE_DICT) and "CONCATENATE" not in str(VALUE_DICT) and "N/A" not in str(VALUE_DICT) and len(str(VALUE_DICT)) < 100:
                                                            for key, val in VALUE_DICT.items():
                                                                if "yes" not in val.lower() and "/" not in val:
                                                                        INTERN_EXTRA_DICT = "\n".join([INTERN_EXTRA_DICT, val])
                                    value = value + 1
                                col = col + 1
                        except:
                            col = col + 1
                            end_col_count = end_col_count + 1
                            if end_col_count == 15:
                                end_of_col = 1
                            if end_of_col:
                                break
                    if str(BACKGROUND_COLOR) != "{'red': 1, 'blue': 1}":
                        data_avilable = 1
                        break
                else:
                    row = row + 1
            except:
                row = row + 1
                end_row_count = end_row_count + 1
                if end_row_count == 10:
                    end_of_sheet = 1
                if end_of_sheet:
                    break
        if data_avilable:
            break

    return INTERN_NAME_DICT, INTERN_NUM_DICT, INTERN_EMAIL_DICT, INTERN_EXTRA_DICT, BACKGROUND_COLOR, data_avilable


@lru_cache
def status_check(color):
    magenta, white, green, red, yellow, cyan, status = "{'red': 1, 'blue': 1}", "{'red': 1, 'green': 1, 'blue': 1}", "{'green': 1}", "{'red': 1}", "{'red': 1, 'green': 1}", "{'green': 1, 'blue': 1}", ""

    if color == magenta:
        status = "working under other HR"
    elif color == white:
        status = "Not Interviewed"
    elif color == green:
        status = "selected"
    elif color == red:
        status = "Not Interested"
    elif color == yellow:
        status = "Interviewed"
    elif color == cyan:
        status = "Call didn't connect"
    return status


root = Tk()
root.geometry("1500x1000")
root.title("SMIDST - Social Media Intern's Data Searching Tool")

# ------bgx = "" will give error
user_inp, INFO_TO_COPY, draft_text, got_email_value, bgx = None, "", "", "", "white"


def copy_button():
    global root
    global INFO_TO_COPY
    root.clipboard_clear()
    root.clipboard_append(INFO_TO_COPY)
    INFO_TO_COPY = ""


def draft_copy():
    global draft_text
    global INFO_TO_COPY
    INFO_TO_COPY = draft_text
    copy_button()


def email_copy():
    global got_email_value
    global INFO_TO_COPY
    INFO_TO_COPY = got_email_value
    copy_button()


def show_details():
    global user_number
    global user_name
    global user_inp
    global HR_contact
    global details
    global url
    global INFO_TO_COPY
    global bgx
    global draft_text
    global got_email_value
    global no_internet
    show = Frame(bg="#3b3a39")
    show.place(width="1500", height="1000",
               rely=0.5, relx=0.5, anchor="center")

    # calling main function with variable details
    Name, Contact, got_email, got_email_value, Extra_details, draft_text, details = "", "", 0, "", "", "", None

    try:
        user_number = int(user_inp.get())
        user_number = user_inp.get()
    except:
        user_name = user_inp.get()

    if user_name != None:
        details = main_byName(user_name)
    elif user_number != None:
        details = main(user_number)

    if no_internet:
        No_net_label = Label(show, text=f"Please check your internet connection !!!",
                             bg="#3b3a39", fg="#0289f7", font=("Eras Demi ITC", 25))
        No_net_label.place(x=440, y=350)
        backbtn = Button(show, text="Back", width=13, height=2, bg="#2982ff",
                     fg="white", font=("Eras Demi ITC", 12), command=Home)
        backbtn.place(x=50, y=100)
        return 0

    if details[0] != "":
        for key, value in details[0].items():
            Name = value
    if details[1] != "":
        for key, value in details[1].items():
            Contact = str(value)
    if details[2] != "":
        for key, value in details[2].items():
            if "@" in value:
                got_email = 1
                got_email_value = value
    if details[3] != "":
        Extra_details = details[3]

    status = status_check(str(details[4]))

    if status == "selected":
        bgx = "#15bd02"
    elif status == "Not Interested":
        bgx = "#ff0000"
    elif status == "working under other HR":
        bgx = "#ff00bf"
    elif status == "Not Interviewed":
        bgx = "#ffffff"
    elif status == "Interviewed":
        bgx = "#f7ca02"
    elif status == "Call didn't connect":
        bgx = "#02f7de"

    HR_details = HR_search()

    HR_contact, HR_bgx, HR_data_avilable = HR_details[0], HR_details[1], HR_details[2]

    backbtn = Button(show, text="Back", width=13, height=2, bg="#2982ff",
                     fg="white", font=("Eras Demi ITC", 12), command=Home)
    backbtn.place(x=50, y=100)
    if details[5]:

        title = Label(show, text=Name, font=(
            "Algerian", 50), bg="#3b3a39", fg=bgx)
        title.pack(side="top", pady=90, anchor="n")

        contact_label = Label(show, text="Contact Number:",
                              bg="#3b3a39", fg="#0289f7", font=("Eras Demi ITC", 18))
        contact_label.place(x=160, y=200)
        Contact_number = Label(
            show, text=Contact, bg="#3b3a39", fg="white", font=("Eras Demi ITC", 18))
        Contact_number.place(x=160, y=250)

        if len(Contact) > 11:
            rm = len(Contact) - 10
            Contact = Contact[rm:]
        url = f"http://wa.me/91{Contact}"
        WhatsApp_Btn = Button(show, text="WhatsApp", width=17, height=2, bg="#01e675",
                              fg="white", font=("Helvetica", 11), command=openweb)
        WhatsApp_Btn.place(x=160, y=300)

        status_label = Label(show, text=f"Status",
                             bg="#3b3a39", fg="#0289f7", font=("Eras Demi ITC", 18))
        status_label.place(x=160, y=380)
        intern_status = Label(show, text=status, bg="#3b3a39",
                              fg="white", font=("Eras Demi ITC", 18))
        intern_status.place(x=160, y=415)

        Other_details_label = Label(
            show, text="Other Details", bg="#3b3a39", fg="#0289f7", font=("Eras Demi ITC", 18))
        Other_details_label.place(x=160, y=480)
        Other_details = Message(show, text=f"{got_email_value}\n{Extra_details}", bg="#3b3a39", fg="white", font=(
            "Eras Demi ITC", 18), justify=LEFT)
        Other_details.place(x=160, y=530)

        if got_email:
            email_copy_btn = Button(show, text="Copy Email ID", width=10, height=1, bg="#2982ff",
                                    fg="white", font=("Helvetica", 11), command=email_copy)
            email_copy_btn.place(x=350, y=480)

        if status == "Not Interviewed":
            draft_text = f"Hi {Name}\n We are sorry for the inconvenience,\n your HR will contact you within 2-3 days"
        elif status == "Interviewed":
            draft_text = f"Hi {Name}\n As per our records your HR is already in contact with you"
        elif status == "Call didn't connect":
            draft_text = f"Hi {Name}\n We tried to contact you for the telephonic interview but, we were unable to connect you\n possible reasons are,\n 1. You didn't pick up the call\n 2. Do not have active incomming service avilable on your number {Contact} \n 3. Any other technical issue\n *But don't worry your HR will contact you again* "
        if draft_text != "":
            draft_copy_btn = Button(show, text="Copy", width=10, height=1, bg="#2982ff",
                                    fg="white", font=("Helvetica", 11), command=draft_copy)
            draft_copy_btn.place(x=1000, y=200)
            draft_text_label = Label(
                show, text="Draft Text", bg="#3b3a39", fg="#0289f7", font=("Eras Demi ITC", 18))
            draft_text_label.place(x=800, y=200)
            show_draft_text = Message(show, text=draft_text, bg="#3b3a39", fg="white", font=(
                "Eras Demi ITC", 15), justify=LEFT)
            show_draft_text.place(x=800, y=250)

        if HR_data_avilable:

            HR_info_label = Label(
                show, text="HR Information", bg='#3b3a39', fg='#0289f7', font=("Eras Demi ITC", 18))
            HR_info_label.place(x=800, y=510)
            HR_name = Label(show, text=HR_details[3], bg="#3b3a39", fg="white", font=(
                "Eras Demi ITC", 18), justify=LEFT)
            HR_name.place(x=800, y=560)
            HR_contact_show = Label(show, text=HR_contact, bg="#3b3a39", fg="white", font=(
                "Eras Demi ITC", 18), justify=LEFT)
            HR_contact_show.place(x=800, y=610)

            WhatsApp_Btn = Button(show, text="WhatsApp", width=17, height=2, bg="#01e675",
                                  fg="white", font=("Helvetica", 11), command=set_HR_url)
            WhatsApp_Btn.place(x=1000, y=510)
            if HR_bgx == {'red': 1}:
                HR_warning1 = Label(show, text="HR is marked red", bg="#3b3a39", fg=HR_bgx, font=(
                    "Eras Demi ITC", 18), justify=LEFT)
                HR_warning1.place(x=800, y=660)
        else:
            pass

    else:
        No_data_label = Label(show, text=f"No data avilable for '*'",
                              bg="#3b3a39", fg="#0289f7", font=("Eras Demi ITC", 25))
        No_data_label.place(x=460, y=350)


def Home():
    global user_number
    global user_name
    global user_inp
    user_name = None
    user_number = None

    home = Frame(bg="#3b3a39")
    home.place(width="1500", height="1000",
               rely=0.5, relx=0.5, anchor="center")
    title = Label(home, text="SMIDST", font=(
        "Algerian", 50), bg="#3b3a39", fg="white")
    title.pack(side="top", pady=90, anchor="n")

    para1 = Label(home, text="Enter contact number of an intern",
                  bg="#3b3a39", fg="white", font=("Eras Demi ITC", 18))
    para1.place(x=500, y=250)

    user_inp = Entry(home, width=25, font=("Arial", 20))
    user_inp.place(x=500, y=300)

    b1 = Button(home, text="Submit", width=13, height=2,
                bg="#2982ff", fg="white", font=("Eras Demi ITC", 12), command=show_details)
    b1.place(x=500, y=450)


Home()

result_creator()
root.mainloop()

# -------Features added
# -------1. Search by name
# -------2. WhatsApp of HR
# -------3. Faster

# -------Features to add
# -------scroll "other details" section
# -------reload button if "no network"
# -------Read any sheet (itrate key from external file)
# -------Sheet name (location)
