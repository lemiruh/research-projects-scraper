# This Python file uses the following encoding: utf-8
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import xlsxwriter

MAIN_HEADER = ["Order", "Source", "Exercise Year :", "Project Number :", "Project Title(English) :",
               "Principal Investigator(English) :", "Project Status :", "Funding Scheme :", "Project Title(Chinese) :",
               "Principal Investigator(Chinese) :", "Department :", "Institution :", "E-mail Address :", "Tel :",
               "Co - Investigator(s) :", "Panel :", "Subject Area :", "Fund Approved :", "Completion Date :",
               "Abstract as per original application\n(English/Chinese):", "Realisation of objectives:",
               "Major findings and research outcome:",
               "Potential for further development of the research\nand the proposed course of action:",
               "Layman's Summary of\nCompletion Report:", "Completion Report:", "Last Update on :"]
OBJECTIVE_ACHIEVED_HEADER = ["Exercise Year :", "Project Number :", "Principal Investigator(English) :",
                             "Project Title(English) :", "Project Objectives :",
                             "Addressed", "Percentage Achieved"]
RESEARCH_OUTPUT_1_HEADER = ["Exercise Year :", "Project Number :", "Principal Investigator(English) :",
                            "Project Title(English) :", "Year of Publication", "Author(s)", "Title and Journal/Book",
                            "Accessible from Institution Repository"]
RESEARCH_OUTPUT_2_HEADER = ["Exercise Year :", "Project Number :", "Principal Investigator(English) :",
                            "Project Title(English) :", "Month/Year/City", "Title", "Conference Name"]
OTHER_IMPACTS_HEADER = ["Exercise Year :", "Project Number :", "Principal Investigator(English) :",
                        "Project Title(English) :", "Other impact\n(e.g. award of patents or prizes,\ncollaboration with other research institutions,\ntechnology transfer, etc.):"]


def write_sheet_from_sheet_of_table(workbook, sheet, sheet_header, sheet_name):
    worksheet = workbook.add_worksheet(sheet_name)
    for col_num, data in enumerate(sheet_header):
        worksheet.write(0, col_num, data)
    row_num = 0
    for rows in sheet:
        for row in rows:
            for col_num, header_column in enumerate(sheet_header):
                worksheet.write(row_num+1, col_num, row[header_column])
            row_num += 1


def write_sheet_from_sheet_of_row(workbook, sheet, sheet_header, sheet_name):
    worksheet = workbook.add_worksheet(sheet_name)
    for col_num, data in enumerate(sheet_header):
        worksheet.write(0, col_num, data)
    for row_num, row in enumerate(sheet):
        for col_num, header_column in enumerate(sheet_header):
            try:
                worksheet.write(row_num+1, col_num, row[header_column])
            except: # the header wasn't found in the table, write blank. HAS TO BE IMPROVED
                worksheet.write(row_num + 1, col_num, '')


def write_xlsx_file(file_name, main_sheet, objective_achieved_sheet, research_output_1_sheet,
                    research_output_2_sheet, other_impact_sheet):

    workbook = xlsxwriter.Workbook(file_name + '.xlsx')

    # MAIN SHEET
    write_sheet_from_sheet_of_row(workbook, main_sheet, MAIN_HEADER, "Main")
    # OBJECTIVES SHEET
    write_sheet_from_sheet_of_table(workbook, objective_achieved_sheet, OBJECTIVE_ACHIEVED_HEADER, "Objective Achieved")
    # RESEARCH OUTPUT 1 SHEET
    write_sheet_from_sheet_of_table(workbook, research_output_1_sheet, RESEARCH_OUTPUT_1_HEADER, "Research Output1")
    # RESEARCH OUTPUT 2 SHEET
    write_sheet_from_sheet_of_table(workbook, research_output_2_sheet, RESEARCH_OUTPUT_2_HEADER, "Research Output2")
    # OTHER IMPACT SHEET
    write_sheet_from_sheet_of_row(workbook, other_impact_sheet, OTHER_IMPACTS_HEADER, "Other Impacts")

    workbook.close()


def create_objective_achieved_row(target_table_td_tag_elements, main_row):
    new_row = {}
    for i in range(4):
        new_row[OBJECTIVE_ACHIEVED_HEADER[i]] = main_row[OBJECTIVE_ACHIEVED_HEADER[i]]
    new_row["Project Objectives :"] = target_table_td_tag_elements[1].text
    new_row["Addressed"] = target_table_td_tag_elements[2].text
    new_row["Percentage Achieved"] = target_table_td_tag_elements[3].text

    return new_row


def create_research_output_1_row(target_table_td_tag_elements, main_row):
    new_row = {}
    for i in range(4):
        new_row[RESEARCH_OUTPUT_1_HEADER[i]] = main_row[RESEARCH_OUTPUT_1_HEADER[i]]
    new_row["Year of Publication"] = target_table_td_tag_elements[0].text
    new_row["Author(s)"] = target_table_td_tag_elements[1].text
    new_row["Title and Journal/Book"] = target_table_td_tag_elements[2].text
    new_row["Accessible from Institution Repository"] = target_table_td_tag_elements[3].text

    return new_row


def create_research_output_2_row(target_table_td_tag_elements, main_row):
    new_row = {}
    for i in range(4):
        new_row[RESEARCH_OUTPUT_2_HEADER[i]] = main_row[RESEARCH_OUTPUT_2_HEADER[i]]
    new_row["Month/Year/City"] = target_table_td_tag_elements[0].text
    new_row["Title"] = target_table_td_tag_elements[1].text
    new_row["Conference Name"] = target_table_td_tag_elements[2].text

    return new_row


def retrieve_data_of_one_project(driver):
    #the table we want
    table = driver.find_elements(By.XPATH, "//body/table/tbody")[3]
    tr_tag_elements = table.find_elements(By.XPATH, "./*")
    main_row = {}
    other_impacts_row = {}
    objective_achieved_rows = []
    research_output_1_rows = []
    research_output_2_rows = []

    header_of_row_with_table = ["Summary of objectives addressed:",
                                "Peer-reviewed journal publication(s)\narising directly from this research project :\n(* denotes the corresponding author)",
                                "Recognized international conference(s)\nin which paper(s) related to this research\nproject was/were delivered :"]

    for tr_tag_element in tr_tag_elements:
        td_tag_elements = tr_tag_element.find_elements(By.XPATH, "./*")

        if len(td_tag_elements) > 1:
            row_header = td_tag_elements[0].text

            # BUILD MAIN ROW
            if MAIN_HEADER.__contains__(row_header):
                main_row[row_header] = td_tag_elements[1].text

            # BUILD OTHER IMPACTS
            elif OTHER_IMPACTS_HEADER.__contains__(row_header):
                if td_tag_elements[1].text != '':
                    for i in range(4):
                        other_impacts_row[OTHER_IMPACTS_HEADER[i]] = main_row[OTHER_IMPACTS_HEADER[i]]
                    other_impacts_row[row_header] = td_tag_elements[1].text

            # IF THE CONTENT OF THE ROW IS ANOTHER TABLE
            elif header_of_row_with_table.__contains__(row_header):
                target_table = td_tag_elements[1].find_element(By.XPATH, "./*")
                target_table_children = target_table.find_elements(By.XPATH, "./*")
                if len(target_table_children) != 0: # there is data in the table
                    target_table_body = target_table_children[0]
                    target_table_tr_tag_elements = target_table_body.find_elements(By.XPATH, "./*")
                    target_table_tr_tag_elements.pop(0)  # remove the header

                    for target_table_tr_tag_element in target_table_tr_tag_elements:
                        new_row = {}
                        target_table_td_tag_elements = target_table_tr_tag_element.find_elements(By.XPATH, "./*")

                        # BUILD OBJECTIVE ACHIEVED ROWS
                        if row_header == header_of_row_with_table[0]:
                            if len(target_table_td_tag_elements) == 4: # reject case where it's N/A (project number 100210 for example)
                                objective_achieved_rows.append(create_objective_achieved_row(target_table_td_tag_elements, main_row))

                        # BUILD RESEARCH OUTPUT1
                        elif row_header == header_of_row_with_table[1]:
                            research_output_1_rows.append(create_research_output_1_row(target_table_td_tag_elements, main_row))

                        # BUILD RESEARCH OUTPUT2
                        elif row_header == header_of_row_with_table[2]:
                            research_output_2_rows.append(create_research_output_2_row(target_table_td_tag_elements, main_row))

    return main_row, objective_achieved_rows, research_output_1_rows, research_output_2_rows, other_impacts_row


def main(projects_year):
    driver = webdriver.Chrome()
    driver.get("website_link")

    # select the year drop down menu
    drop_down_menu = Select(driver.find_element(By.NAME, "Year"))
    drop_down_menu.select_by_value(projects_year)

    # find the form and submit to display every research paper
    form_search = driver.find_element(By.NAME, "___")
    form_search.submit()

    main_sheet = []
    other_impact_sheet = []
    objective_achieved_sheet = []
    research_output_1_sheet = []
    research_output_2_sheet = []
    current_page_is_the_last = False
    # loop for every page of projects
    index = 0
    while current_page_is_the_last == False and index < 10:
        # loop to go to the projects of the page
        for i in range(len(driver.find_elements(By.NAME, "theSubmit"))):
            # find the buttons to go to projects
            project_buttons = driver.find_elements(By.NAME, "theSubmit")
            project_buttons[i].click()

            main_row, objective_achieved_rows, research_output_1_rows, research_output_2_rows, other_impact_row = retrieve_data_of_one_project(driver)
            main_sheet.append(main_row)
            objective_achieved_sheet.append(objective_achieved_rows)
            research_output_1_sheet.append(research_output_1_rows)
            research_output_2_sheet.append(research_output_2_rows)
            if len(other_impact_row) != 0:
                other_impact_sheet.append(other_impact_row)

            # go to previous page
            project_page_buttons = driver.find_elements(By.TAG_NAME, "input")
            for project_page_button in project_page_buttons:
                if project_page_button.get_attribute("value") == "  Return  ":
                    project_page_button.click()

        # search for the next button
        current_page_is_the_last = True
        navigate_in_pages_buttons = driver.find_elements(By.XPATH, "//a")
        for navigate_in_pages_button in navigate_in_pages_buttons:
            if navigate_in_pages_button.text == "[Next Page]":
                navigate_in_pages_button.click()
                current_page_is_the_last = False
                break

        index += 1

    write_xlsx_file("output_year_" + projects_year, main_sheet, objective_achieved_sheet, research_output_1_sheet, research_output_2_sheet, other_impact_sheet)

main('2020')
# main('2021')
# main('2022')
