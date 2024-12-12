# import xlwings as xw
# import time
# import json
# import sys
# import os
# import logging
# import pygetwindow as pgw
# import pyautogui
# from pynput.mouse import Button, Controller as mouseController
# from pynput.keyboard import Key, Controller as keyboardController
#
# pyautogui.PAUSE = 2.5
# mouse = mouseController()
# keyboard = keyboardController()
#
# coordinateMap = {}
# userInput = {}
#
# def loadConfig():
#     try:
#         if os.path.isfile('../user_input/UserInput.json'):
#             with open('../user_input/UserInput.json', "r") as f_userInput:
#                 global userInput
#                 userInput = json.load(f_userInput)
#         else:
#             logging.info("UserInput.json File not found in /user_input folder")
#             exit()
#     except Exception as ex:
#         exc_type, exc_obj, exc_tb = sys.exc_info()
#         exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
#         print(f"Error{ex} line no. {exc_tb.tb_lineno} file name: {exp_fname}")
#
# def getTestPlanMacro():
#     return userInput["toolsPath"]["TestPlanMacro"]
#
# def getOutputFiles():
#     return userInput["toolsPath"]["OutputFolder"]
#
# def print:
#     return userInput["toolsPath"]["InputFolder"]
#
# def getDownloadFolder():
#     return userInput["toolsPath"]["DownloadFolder"]
#
# def openExcel(book):
#     return xw.Book(book)
#
# def findInputFiles():
#     arr = os.listdir(getInputFolder())
#     PT, CD = "", ""
#     for i in arr:
#         if i.find('ESAD_Task_Tracking_Sheet') != -1 and i.find('~$') == -1:
#             PT = i
#         if i.find('Carnet_de') != -1 and i.find('~$') == -1:
#             CD = i
#     return [PT, CD]
#
# def openTestPlan():
#     PT = findInputFiles()[0]
#     if len(PT) != 0:
#         testPlan = openExcel(print + "\\" + PT)
#         return testPlan
#     else:
#         logging.info("No testplan")
#         return -1
#
# def getDataFromCell(sheet, colRow):
#     return sheet.range(colRow).value
#
# def setDataFromCell(sheet, colRow, value):
#     sheet.range(colRow).value = value
#
# def searchDataInColCache(value, specfCol, keyword, matchCase=False):
#     searchResult = {
#         "count": 0,
#         "cellPositions": [],
#         "cellValue": []
#     }
#     if keyword == "":
#         return searchResult
#     for x, i in enumerate(value):
#         for y, j in enumerate(i):
#             if y == specfCol - 1:
#                 if j is not None:
#                     if matchCase == True:
#                         if keyword.lower() in str(j).lower():
#                             searchResult["count"] = searchResult["count"] + 1
#                             searchResult["cellPositions"].append((x + 1, y + 1))
#                             searchResult["cellValue"].append(j)
#                     else:
#                         if keyword in str(j):
#                             searchResult["count"] = searchResult["count"] + 1
#                             searchResult["cellPositions"].append((x + 1, y + 1))
#                             searchResult["cellValue"].append(j)
#     return searchResult
#
# def rightArrow():
#     keyboard.press(Key.right)
#     keyboard.release(Key.right)
#
# def pressEnter():
#     keyboard.press(Key.enter)
#     keyboard.release(Key.enter)
#
# def excel_popup(windowName):
#     while (True):
#         window_title = pgw.getActiveWindowTitle()
#         global stop_threads
#         excel_windows = pgw.getWindowsWithTitle("Excel")
#         try:
#             for each_excel_window in excel_windows:
#                 if windowName.split('.')[0] in each_excel_window.title:
#                     each_excel_window.minimize()
#                     each_excel_window.maximize()
#                     if not each_excel_window.isActive:
#                         each_excel_window.activate()
#                         break
#                 else:
#                     each_excel_window.minimize()
#         except:
#             print("Exception in excel popup")
#             break
#         if pgw.getActiveWindowTitle() == "Microsoft Excel":
#             time.sleep(1)
#             rightArrow()
#             time.sleep(1)
#             pressEnter()
#         active_window = pgw.getActiveWindow()
#         if active_window is not None:
#             active_window.minimize()
#         if stop_threads:
#             break
#
# def searchDataInExcelCache(value, keyword):
#     searchResult = {
#         "count": 0,
#         "cellPositions": [],
#         "cellValue": []
#     }
#     if keyword == "":
#         return searchResult
#
#     for x, i in enumerate(value):
#         for y, j in enumerate(i):
#             if j is not None:
#                 if keyword in str(j):
#                     searchResult["count"] = searchResult["count"] + 1
#                     searchResult["cellPositions"].append((x + 1, y + 1))
#                     searchResult["cellValue"].append(j)
#
#     return searchResult
#
# def main(tab):
#     FELP_Nos, Components, Task_Types, Simple_Complexitys, Medium_complexitys, New_implementations = [], [], [], [], [], []
#     SWC_versions, Task_received_dates, PSA_Review_Deadlines, Task_Final_Submissions = [], [], [], []
#     PSA_Final_Submission_Date_requesteds, MCDCs, Qualtiys, misras = [], [], [], []
#     try:
#         start = time.time()
#         tpBook = openTestPlan()
#         sheet = tpBook.sheets[tab]
#         maxrow = tpBook.sheets[sheet].range('A' + str(tpBook.sheets[sheet].cells.last_cell.row)).end('up').row
#         Check_list_sTr_sheet = tpBook.sheets[sheet]
#         sheet_value = Check_list_sTr_sheet.used_range.value
#         config_name = searchDataInExcelCache(sheet_value, 'Task Status')
#         row, col = config_name['cellPositions'][0]
#         deliver = searchDataInColCache(sheet_value, col, 'Delivered')
#
#         for i in deliver['cellPositions']:
#             row, col = i
#             FELP_No = getDataFromCell(tpBook.sheets[sheet], (row, col-19))
#             Component = getDataFromCell(tpBook.sheets[sheet], (row, col-17))
#             Task_Type = getDataFromCell(tpBook.sheets[sheet], (row, col-15))
#             Simple_Complexity = getDataFromCell(tpBook.sheets[sheet], (row, col-13))
#             if Simple_Complexity == 'None':
#                 Simple_Complexity = 0
#             Medium_complexity = getDataFromCell(tpBook.sheets[sheet], (row, col-12))
#             if Medium_complexity == 'None':
#                 Medium_complexity = 0
#             New_implementation = getDataFromCell(tpBook.sheets[sheet], (row, col-11))
#             if New_implementation == 'None':
#                 New_implementation = 0
#             SWC_version = getDataFromCell(tpBook.sheets[sheet], (row, col - 9))
#             Task_received_date = getDataFromCell(tpBook.sheets[sheet], (row, col - 7))
#             PSA_Review_Deadline = getDataFromCell(tpBook.sheets[sheet], (row, col + 4))
#             Task_Final_Submission = getDataFromCell(tpBook.sheets[sheet], (row, col + 6))
#             PSA_Final_Submission_Date_requested = getDataFromCell(tpBook.sheets[sheet], (row, col + 7))
#             MCDC = getDataFromCell(tpBook.sheets[sheet], (row, col + 9))
#             Qualtiy = getDataFromCell(tpBook.sheets[sheet], (row, col + 12))
#             misra = getDataFromCell(tpBook.sheets[sheet], (row, col + 14))
#
#             FELP_Nos.append(FELP_No), Components.append(Component), Task_Types.append(Task_Type), Simple_Complexitys.append(Simple_Complexity), Medium_complexitys.append(Medium_complexity), New_implementations.append(New_implementation)
#             SWC_versions.append(SWC_version), Task_received_dates.append(Task_received_date), PSA_Review_Deadlines.append(PSA_Review_Deadline), Task_Final_Submissions.append(Task_Final_Submission)
#             PSA_Final_Submission_Date_requesteds.append(PSA_Final_Submission_Date_requested), MCDCs.append(MCDC), Qualtiys.append(Qualtiy), misras.append(misra)
#
#         hsg = list(zip(FELP_Nos, Components, Task_Types, Simple_Complexitys, Medium_complexitys, New_implementations, SWC_versions, Task_received_dates, PSA_Review_Deadlines, Task_Final_Submissions, PSA_Final_Submission_Date_requesteds, MCDCs, Qualtiys, misras))
#         path = print + "\\" + findInputFiles()[1]
#         Carnet_de_Book = openExcel(path)
#         Carnet_de_Book.activate()
#         maxrow = Carnet_de_Book.sheets['Carnet de commande'].range('B' + str(Carnet_de_Book.sheets['Carnet de commande'].cells.last_cell.row)).end('up').row
#
#         all_values = Carnet_de_Book.sheets['Carnet de commande'].range('A1:BK' + str(maxrow)).value
#         valid_row = searchDataInColCache(all_values, 20, 'Sent')
#         win_name = path.split('\\')[-1]
#         stop_threads = False
#         popup = threading.Thread(target=excel_popup, args=(win_name,))
#         popup.start()
#
#         time.sleep(7)
#         for val in valid_row['cellPositions']:
#             row, col = val
#             Task_Submission_PSA = getDataFromCell(Carnet_de_Book.sheets['Carnet de commande'], (row, 66))
#             if Task_Submission_PSA is None:
#                 row, col = val
#                 misra, qualtiy, task_Type, Component = "", "", "", ""
#                 PSA_review_deadline, PSA_Review_Deadline, Task_Final_Submission, Task_Final_Submission_requested = "", "", "", ""
#                 Simple_Complexity, Medium_complexity, New_implementation = "", "", ""
#
#                 felpNo = getDataFromCell(Carnet_de_Book.sheets['Carnet de commande'], (row, col - 18))
#
#                 for tup in hsg:
#                     if tup[0] == felpNo:
#                         Component, task_Type, Simple_Complexity, Medium_complexity, New_implementation, swc_version = tup[1], tup[2], tup[3], tup[4], tup[5], tup[6]
#                         task_received_date, PSA_review_deadline, Task_Final_Submission, Task_Final_Submission_requested, MCDC, qualtiy, misra = tup[7], tup[8], tup[9], tup[10], tup[11], tup[12], tup[13]
#                         break
#
#                 setDataFromCell(Carnet_de_Book.sheets['Carnet de commande'], (row, col - 14), Component)
#                 setDataFromCell(Carnet_de_Book.sheets['Carnet de commande'], (row, col - 13), task_Type)
#                 setDataFromCell(Carnet_de_Book.sheets['Carnet de commande'], (row, col - 12), Simple_Complexity)
#                 setDataFromCell(Carnet_de_Book.sheets['Carnet de commande'], (row, col - 11), Medium_complexity)
#                 setDataFromCell(Carnet_de_Book.sheets['Carnet de commande'], (row, col - 10), New_implementation)
#                 setDataFromCell(Carnet_de_Book.sheets['Carnet de commande'], (row, col - 8), swc_version)
#                 setDataFromCell(Carnet_de_Book.sheets['Carnet de commande'], (row, col - 7), task_received_date)
#                 setDataFromCell(Carnet_de_Book.sheets['Carnet de commande'], (row, col + 4), PSA_review_deadline)
#                 setDataFromCell(Carnet_de_Book.sheets['Carnet de commande'], (row, col + 6), Task_Final_Submission)
#                 setDataFromCell(Carnet_de_Book.sheets['Carnet de commande'], (row, col + 7), Task_Final_Submission_requested)
#                 setDataFromCell(Carnet_de_Book.sheets['Carnet de commande'], (row, col + 9), MCDC)
#                 setDataFromCell(Carnet_de_Book.sheets['Carnet de commande'], (row, col + 12), qualtiy)
#                 setDataFromCell(Carnet_de_Book.sheets['Carnet de commande'], (row, col + 14), misra)
#
#                 keyboard.press(Key.enter)
#                 keyboard.release(Key.enter)
#
#         time.sleep(3)
#         stop_threads = True
#         popup.join()
#         Carnet_de_Book.save()
#         Carnet_de_Book.close()
#
#     except Exception as ex:
#         exc_type, exc_obj, exc_tb = sys.exc_info()
#         exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
#         print(f"Error{ex} line no. {exc_tb.tb_lineno} file name: {exp_fname}")
#
# if __name__ == '__main__':
#     loadConfig()
#     TestPlanMacro = getTestPlanMacro()
#     xl = xw.App(visible=True, add_book=False)
#     tab = sys.argv[1]
#     main(tab)
#     xl.quit()



from datetime import datetime, timedelta
import xlwings as xw
import time
import json
import sys
import os
import pygetwindow as pgw
import pyautogui
from pynput.mouse import Button, Controller as mouseController
from pynput.keyboard import Key, Controller as keyboardController
import logging
pyautogui.PAUSE = 2.5
mouse = mouseController()
keyboard = keyboardController()



def openExcel(book):
    return xw.Book(book)


def findInputFiles(path):
    arr = os.listdir(path)
    PT, CD = "", ""
    for i in arr:
        if 'ESAD_Task_Tracking_Sheet' in i and '~$' not in i:
            PT = i
        if 'Carnet_de' in i and '~$' not in i:
            CD = i
    return [PT, CD]


def openTestPlan(path):
    PT = findInputFiles(path)[0]
    if len(PT) != 0:
        testPlan = openExcel(os.path.join(path, PT))
        return testPlan
    else:
        logging.info("No testplan")
        return -1


def getDataFromCell(sheet, colRow):
    return sheet.range(colRow).value


def setDataFromCell(sheet, colRow, value):
    sheet.range(colRow).value = value


def searchDataInColCache(value, specfCol, keyword, matchCase=False):
    searchResult = {
        "count": 0,
        "cellPositions": [],
        "cellValue": []
    }
    if keyword == "":
        return searchResult
    for x, i in enumerate(value):
        for y, j in enumerate(i):
            if y == specfCol - 1:
                if j is not None:
                    # logging.info("j --- ", j)
                    if matchCase == True:
                        if keyword.lower() in str(j).lower():
                            searchResult["count"] = searchResult["count"] + 1
                            searchResult["cellPositions"].append((x + 1, y + 1))
                            searchResult["cellValue"].append(j)
                    else:
                        if keyword in str(j):
                            searchResult["count"] = searchResult["count"] + 1
                            searchResult["cellPositions"].append((x + 1, y + 1))
                            searchResult["cellValue"].append(j)
    return searchResult


def rightArrow():
    # Press and release enter
    keyboard.press(Key.right)
    keyboard.release(Key.right)


def pressEnter():
    # Press and release enter
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)


def excel_popup(windowName):
    while (True):
        window_title = pgw.getActiveWindowTitle()
        global stop_threads
        excel_windows = pgw.getWindowsWithTitle("Excel")
        try:
            for each_excel_window in excel_windows:
                if windowName.split('.')[0] in each_excel_window.title:
                    each_excel_window.minimize()
                    each_excel_window.maximize()
                    if each_excel_window.isActive == False:
                        each_excel_window.activate()
                        break
                else:
                    each_excel_window.minimize()
        except:
            logging.info("Exception in excel popup")
            break
        if pgw.getActiveWindowTitle() == "Microsoft Excel":
            time.sleep(1)
            rightArrow()
            time.sleep(1)
            pressEnter()
        active_window = pgw.getActiveWindow()
        if active_window is not None:
            active_window.minimize()
        if stop_threads:
            break


def searchDataInExcelCache(value, keyword):
    # value = sheet.used_range.value
    searchResult = {
        "count": 0,
        "cellPositions": [],
        "cellValue": []
    }
    # x is the index of column
    # i is the value of column
    # y is the index of row
    # j is the value of cell
    if keyword=="":
        return searchResult

    for x, i in enumerate(value):
        for y, j in enumerate(i):
            if j is not None:
                if keyword in str(j):
                    searchResult["count"] = searchResult["count"] + 1
                    searchResult["cellPositions"].append((x + 1, y + 1))
                    searchResult["cellValue"].append(j)

    return searchResult


def get_last_n_months(month_number, start_year, n):
    """ Get the last n months including a given month and year, in ascending order as month names with years. """
    month_names = ["Jan", "Feb", "March", "April", "May", "June",
                   "July", "Aug", "Sep", "Oct", "Nov", "Dec"]
    months = []
    current_month = month_number - 1  # Adjust for zero-based index
    current_year = start_year

    # Collect the months including the given month
    for _ in range(n):
        months.append(f"{month_names[current_month]} {current_year}")
        if current_month == 0:
            current_month = 11
            current_year -= 1
        else:
            current_month -= 1

    # Sort months in ascending order
    months.reverse()

    return months


# Dictionary mapping month names to their respective numbers
month_mapping = {
    'Jan': 1,
    'Feb': 2,
    'Mar': 3,
    'Apr': 4,
    'May': 5,
    'June': 6,
    'July': 7,
    'Aug': 8,
    'Sep': 9,
    'Oct': 10,
    'Nov': 11,
    'Dec': 12
}

def get_month_number(month_name):
    # Lookup the month name in the dictionary and return the corresponding number
    return month_mapping.get(month_name, "Invalid month name")


def main(tab, path):
    FELP_Nos, Components, Task_Types, Simple_Complexitys, Medium_complexitys, New_implementations = [], [], [], [], [], []
    SWC_versions, Task_received_dates, PSA_Review_Deadlines, Task_Submission_PSAs, Task_Final_Submissions = [], [], [], [], []
    PSA_Final_Submission_Date_requesteds, MCDCs, Qualtiys, misras = [], [], [], []

    try:
        start = time.time()
        print("path------------>", path)
        tpBook = openTestPlan(path)
        month_name, year = tab.split()
        month_number = get_month_number(month_name)
        year = int(year)  # Convert year to integer
        print("month_number---------->", month_number)
        months = get_last_n_months(month_number, year, n=5)
        print("Months to check:", months)
        for month in months:
            print("month--------->", month)
            try:
                sheet = tpBook.sheets[month]
                print("sheet------->", sheet)
                maxrow = tpBook.sheets[sheet].range('A' + str(tpBook.sheets[sheet].cells.last_cell.row)).end('up').row
                logging.info("maxrow-------------->", maxrow)
                Check_list_sTr_sheet = tpBook.sheets[sheet]
                sheet_value = Check_list_sTr_sheet.used_range.value
                # config_name = searchDataInColCache(sheet_value, 1, 'Task Status')
                config_name = searchDataInExcelCache(sheet_value, 'Task Status')
                logging.info("config_name------------->", config_name)
                row, col = config_name['cellPositions'][0]
                logging.info("row, col------->", row, col)
                deliver = searchDataInColCache(sheet_value, col, 'Delivered')
                logging.info("deliver-------->", deliver)
                for i in deliver['cellPositions']:
                    row, col = i
                    # logging.info("row, col--------->", row, col)
                    Task_Final_Submission = getDataFromCell(tpBook.sheets[sheet], (row, col + 6))
                    if isinstance(Task_Final_Submission, datetime):
                        # Check if the month is May
                        if Task_Final_Submission.month == month_number:
                            logging.info("month_numbermonth_numbermonth_number--------------->", month_number)
                            # Print row and col if the condition is met
                            logging.info("Searching Month is Present, row and col are:", row, col)
                            FELP_No = getDataFromCell(tpBook.sheets[sheet], (row, col - 19))
                            Component = getDataFromCell(tpBook.sheets[sheet], (row, col - 17))
                            Task_Type = getDataFromCell(tpBook.sheets[sheet], (row, col - 15))
                            Simple_Complexity = getDataFromCell(tpBook.sheets[sheet], (row, col - 13))
                            if Simple_Complexity == 'None':
                                Simple_Complexity = 0
                            Medium_complexity = getDataFromCell(tpBook.sheets[sheet], (row, col - 12))
                            if Medium_complexity == 'None':
                                Medium_complexity = 0
                            New_implementation = getDataFromCell(tpBook.sheets[sheet], (row, col - 11))
                            if New_implementation == 'None':
                                New_implementation = 0
                            SWC_version = getDataFromCell(tpBook.sheets[sheet], (row, col - 9))
                            Task_received_date = getDataFromCell(tpBook.sheets[sheet], (row, col - 7))
                            PSA_Review_Deadline = getDataFromCell(tpBook.sheets[sheet], (row, col + 4))
                            Task_Submission_PSA = getDataFromCell(tpBook.sheets[sheet], (row, col + 5))
                            Task_Final_Submission = getDataFromCell(tpBook.sheets[sheet], (row, col + 6))
                            PSA_Final_Submission_Date_requested = getDataFromCell(tpBook.sheets[sheet], (row, col + 7))
                            MCDC = getDataFromCell(tpBook.sheets[sheet], (row, col + 9))
                            Qualtiy = getDataFromCell(tpBook.sheets[sheet], (row, col + 12))
                            misra = getDataFromCell(tpBook.sheets[sheet], (row, col + 14))

                            logging.info(
                                "FELP_No,Component,Task_Type,Simple_Complexity,Medium_complexity,New_implementation------->",
                                FELP_No, Component, Task_Type, Simple_Complexity, Medium_complexity, New_implementation)
                            logging.info(
                                "SWC_version,Task_received_date,PSA_Review_Deadline,Task_Submission_PSA,Task_Final_Submission,PSA_Final_Submission_Date_requested------->",
                                SWC_version, Task_received_date, PSA_Review_Deadline, Task_Submission_PSA,
                                Task_Final_Submission, PSA_Final_Submission_Date_requested)
                            logging.info("MCDC,Qualtiy,misra------->", MCDC, Qualtiy, misra)

                            FELP_Nos.append(FELP_No), Components.append(Component), Task_Types.append(
                                Task_Type), Simple_Complexitys.append(Simple_Complexity), Medium_complexitys.append(
                                Medium_complexity), New_implementations.append(New_implementation)
                            SWC_versions.append(SWC_version), Task_received_dates.append(
                                Task_received_date), PSA_Review_Deadlines.append(
                                PSA_Review_Deadline), Task_Submission_PSAs.append(
                                Task_Submission_PSA), Task_Final_Submissions.append(Task_Final_Submission)
                            PSA_Final_Submission_Date_requesteds.append(PSA_Final_Submission_Date_requested), MCDCs.append(
                                MCDC), Qualtiys.append(Qualtiy), misras.append(misra)

                            logging.info(
                                "FELP_Nos, Components, Task_Types, Simple_Complexitys, Medium_complexitys, New_implementations------->",
                                FELP_Nos, Components, Task_Types, Simple_Complexitys, Medium_complexitys, New_implementations)
                            logging.info(
                                "SWC_versions, Task_received_dates, PSA_Review_Deadlines, Task_Submission_PSAs, Task_Final_Submissions------->",
                                SWC_versions, Task_received_dates, PSA_Review_Deadlines, Task_Submission_PSAs,
                                Task_Final_Submissions)
                            logging.info("PSA_Final_Submission_Date_requesteds, MCDCs, Qualtiys, misras------->",
                                         PSA_Final_Submission_Date_requesteds, MCDCs, Qualtiys, misras)
            except:
                print("User given sheet is not present in the Workbook")
        hsg = list(zip(FELP_Nos, Components, Task_Types, Simple_Complexitys, Medium_complexitys, New_implementations,
                       SWC_versions, Task_received_dates, PSA_Review_Deadlines, Task_Submission_PSAs,
                       Task_Final_Submissions, PSA_Final_Submission_Date_requesteds, MCDCs, Qualtiys, misras))
        logging.info("hsg--------->", hsg)
        path = os.path.join(path, findInputFiles(path)[1])
        print("path ---->", path)
        Carnet_de_Book = openExcel(path)
        Carnet_de_Book.activate()
        maxrow = Carnet_de_Book.sheets['Carnet de commande'].range(
            'B' + str(Carnet_de_Book.sheets['Carnet de commande'].cells.last_cell.row)).end('up').row
        logging.info("Carnet de commande maxrow-------------->", maxrow)
        sheet = Carnet_de_Book.sheets['Carnet de commande']
        sheet_value = sheet.used_range.value
        config_name = searchDataInExcelCache(sheet_value, 'Version de la générique')
        logging.info("config_name------------->", config_name)
        row, col = config_name['cellPositions'][0]
        logging.info("row, col------->", row, col)
        sheet = Carnet_de_Book.sheets['Carnet de commande']
        row_index = maxrow + 1  # Assuming maxrow is defined somewhere in your code
        for FELP_No, Component, Task_Type, Simple_Complexity, Medium_complexity, New_implementation, SWC_version, Task_received_date, PSA_Review_Deadline, Task_Submission_PSA, Task_Final_Submission, PSA_Final_Submission_Date_requested, MCDC, Qualtiy, misra in hsg:
            setDataFromCell(sheet, (row_index, col), SWC_version)
            setDataFromCell(sheet, (row_index, col + 1), 'FELP_' + str(int(FELP_No)))
            setDataFromCell(sheet, (row_index, col + 2), Component)
            if Task_Type == 'Branch':
                Task_Type = 'Correction Anomalie'
            if Task_Type == 'Archi':
                Task_Type = 'ARCHI Delivery'
            if Task_Type == 'Trabilite':
                Task_Type = ''
            setDataFromCell(sheet, (row_index, col + 3), Task_Type)
            setDataFromCell(sheet, (row_index, col + 4), Task_received_date)
            setDataFromCell(sheet, (row_index, col + 5), Task_received_date)
            # setDataFromCell(sheet, (row_index, col+6), )
            setDataFromCell(sheet, (row_index, col + 7), 'On Time')
            setDataFromCell(sheet, (row_index, col + 8), PSA_Review_Deadline)
            setDataFromCell(sheet, (row_index, col + 9), 'No')
            # setDataFromCell(sheet, (row_index, col+10), )
            # setDataFromCell(sheet, (row_index, col+11), )
            setDataFromCell(sheet, (row_index, col + 12), '100%')
            setDataFromCell(sheet, (row_index, col + 13), 'livraison officielle faite')
            setDataFromCell(sheet, (row_index, col + 14), Task_Submission_PSA)
            setDataFromCell(sheet, (row_index, col + 16), PSA_Final_Submission_Date_requested)
            setDataFromCell(sheet, (row_index, col + 17), Task_Final_Submission)
            setDataFromCell(sheet, (row_index, col + 29), Simple_Complexity)
            setDataFromCell(sheet, (row_index, col + 30), Medium_complexity)
            setDataFromCell(sheet, (row_index, col + 31), New_implementation)
            setDataFromCell(sheet, (row_index, col + 32), '0')
            # AI below line
            # setDataFromCell(sheet, (row_index, col+32), )
            setDataFromCell(sheet, (row_index, col + 34), 'Manual_Code_UT')
            setDataFromCell(sheet, (row_index, col + 40), MCDC)
            setDataFromCell(sheet, (row_index, col + 41), Qualtiy)
            setDataFromCell(sheet, (row_index, col + 42), misra)

            row_index += 1  # Move to the next row for the next iteration

        Carnet_de_Book.save()
        Carnet_de_Book.close()
        end1 = time.time()
        print("\nexecution time " + str(end1 - start))
        tpBook.close()
    except Exception as ex:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        exp_fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(f"\nerror{ex} line no. {exc_tb.tb_lineno} file name: {exp_fname}")

def open_excel_with_keyword(directory, keyword):
    for filename in os.listdir(directory):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            if keyword.lower() in filename.lower():
                excel_path = os.path.join(directory, filename)
                wb = openExcel(excel_path)
                return wb
    return None


if __name__ == "__main__":
    # loadConfig()
    path = input("Enter the input files folder path: ")
    tab = input("Enter the tab month: ")
    main(tab, path)