""" 
Uses Python 2.7 

This module extracts specific information from the given PDF files
Method:
 - Get the settings information, including file names & paths
 - Extract all data from PDF Files
 - Convert given XLSX to CSV
 - Add data to CSV
 - Save the CSV file as XLSX
"""

import pandas as pd
import slate
import csv
import datetime
import os

from pandas.io.excel import ExcelWriter

settings = {
    "xlsx_filename": "BeneFix Small Group Plans.xlsx",
    "xlsx_sheetname": "Blank Upload Template",
    "pdfs": [
        "para01.pdf",
        "para02.pdf",
        "para03.pdf",
        "para05.pdf",
        "para06.pdf",
        "para07.pdf",
        "para08.pdf",
        "para09.pdf"
    ],
}

# Credit: rogerallen at https://gist.github.com/rogerallen/1583593
us_state_abbrev = {
    'Alabama': 'AL',
    'Alaska': 'AK',
    'Arizona': 'AZ',
    'Arkansas': 'AR',
    'California': 'CA',
    'Colorado': 'CO',
    'Connecticut': 'CT',
    'Delaware': 'DE',
    'Florida': 'FL',
    'Georgia': 'GA',
    'Hawaii': 'HI',
    'Idaho': 'ID',
    'Illinois': 'IL',
    'Indiana': 'IN',
    'Iowa': 'IA',
    'Kansas': 'KS',
    'Kentucky': 'KY',
    'Louisiana': 'LA',
    'Maine': 'ME',
    'Maryland': 'MD',
    'Massachusetts': 'MA',
    'Michigan': 'MI',
    'Minnesota': 'MN',
    'Mississippi': 'MS',
    'Missouri': 'MO',
    'Montana': 'MT',
    'Nebraska': 'NE',
    'Nevada': 'NV',
    'New Hampshire': 'NH',
    'New Jersey': 'NJ',
    'New Mexico': 'NM',
    'New York': 'NY',
    'North Carolina': 'NC',
    'North Dakota': 'ND',
    'Ohio': 'OH',
    'Oklahoma': 'OK',
    'Oregon': 'OR',
    'Pennsylvania': 'PA',
    'Rhode Island': 'RI',
    'South Carolina': 'SC',
    'South Dakota': 'SD',
    'Tennessee': 'TN',
    'Texas': 'TX',
    'Utah': 'UT',
    'Vermont': 'VT',
    'Virginia': 'VA',
    'Washington': 'WA',
    'West Virginia': 'WV',
    'Wisconsin': 'WI',
    'Wyoming': 'WY',
}

def log(string):
    """
    Handles all logging
    """
    time = datetime.datetime.now()
    print("[{0}]: {1}".format(time, string))

def error(string):
    """
    Handles all error messages
    """
    time = datetime.datetime.now()
    print("[{0}] ==== ERROR: ====> {1}".format(time, string))

def parse_PDF(filename):
    """
    Returns the contents of the given PDF file
    filename: String
    """
    try: 
        with open(filename, 'rb') as f:
            doc = slate.PDF(f)
    except IOError as error:
        log("There has been an error with opening file: " + str(filename))
        log(str(error))
        return None

    return [str(page) for page in doc]

def get_valid_dates(page):
    """
    Returns start_date, end_date
    page: List representation of the page - delimiter is a new line
    
    Note: page[0] = "Valid for Effective Dates: MM/DD/YYYY - MM/DD/YYYY"
    """
    line = page[0]
    line = line.split(":")[1]   # Remove everything to the left of the colon
    line = "".join(line.split()) # Remove all whitespace
    line = line.split("-")      # Split into left date and right date round the "-"

    start_date = line[0]
    end_date   = line[1]
    
    return start_date, end_date

def get_product_name(page):
    """
    Returns product_name
    page: List representation of the page - delimiter is a new line
    """
    return page[14]

def get_state(page):
    """
    Returns state in two-letter capitalized form
    page: List representation of the page - delimiter is a new line
    """
    state_name = page[2]
    state_name = state_name[0].upper() + state_name[1:].lower() # Fix casing to match dict

    return us_state_abbrev[state_name]

def get_group_rating_area(page):
    """
    Returns group_rating_area
    page: List representation of the page - delimiter is a new line
    """
    return page[6][:len(page[6]) - 2]

def get_prices(page):
    """
    Returns a list with the prices in the given order:
    zero_eighteen
    nineteen_twenty	
    twenty_one
    twenty_two
    ...
    sixty_three
    sixty_four
    sixty_five_plus
    
    page: List representation of the page - delimiter is a new line
    """
    prices = []

    prices.append(float(page[37]))    # 0-18
    prices.append(float(page[37]))    # 19-20

    for index in range(38, 52):   # 21-34
        prices.append(float(page[index]))

    for index in range(73, 88):   # 35-49
        prices.append(float(page[index]))

    for index in range(110, 125): # 50-64
        prices.append(float(page[index]))

    prices.append(float(page[124]))   # 65+

    return prices

def add_line_to_csv(csv, values):
    """
    Adds a line to the given csv
    csv: python file object, opened with open(filename, "a")
    values: list of values to add
    """
    new_line = ""
    for value in values:
        new_line = new_line + "," + str(value)
    new_line = new_line[1:]

    csv.write(new_line) 
    csv.write("\n")

def xlsx_to_csv(xlsx_filename, xlsx_sheetname):
    """
    Creates a new CSV file with the contents of the given .XLSX file
    Returns the location of the newly created CSV

    xlsx_filename: Path to XLSX File to convert
    xlsx_sheetname: Sheet from which to read
    """
    csv_filename = xlsx_filename.split(".")[0] + '.csv'
    data_xls = pd.read_excel(xlsx_filename, xlsx_sheetname, index=False)
    data_xls.to_csv(csv_filename, encoding='utf-8', index=False)

    return csv_filename

def csv_to_xlsx(csv_filename, xlsx_filename, xlsx_sheetname):
    """
    Write into XLSX file with the contents of the given CSV File
    """
    excel_writer = pd.ExcelWriter(xlsx_filename, engine='xlsxwriter')
    data = pd.read_csv(csv_filename)

    data.to_excel(excel_writer, xlsx_sheetname, index=False)
    excel_writer.save()

def main(settings):
    """
    Executes the script with the given settings
    settings: Object
    """
    if settings == None:
        log("Settings must be provided to main()")
        return 0
    
    lines = []
    for pdf_filename in settings["pdfs"]:
        log("Beginning to parse " + str(pdf_filename))
        pages = parse_PDF(pdf_filename)
        for page in pages:
            page = page.split("\n")
            new_line = []
            new_line = new_line + [get_valid_dates(page)[0], get_valid_dates(page)[1]]
            new_line = new_line + [get_product_name(page)]
            new_line = new_line + [get_state(page)]
            new_line = new_line + [get_group_rating_area(page)]
            new_line = new_line + get_prices(page)
            lines.append(new_line)
        log("Finished parsing " + str(pdf_filename))

    log("Creating temporary CSV")
    csv_filename = xlsx_to_csv(settings["xlsx_filename"], settings["xlsx_sheetname"])

    log("Writing to CSV")
    with open(csv_filename, 'a') as csv:
        for line in lines:
            add_line_to_csv(csv, line)
    
    log("Converting from CSV to XLSX")
    csv_to_xlsx(csv_filename, settings["xlsx_filename"], settings["xlsx_sheetname"])
    os.remove(csv_filename)

main(settings)
