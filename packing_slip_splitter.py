import PyPDF2
from PyPDF2 import PdfFileReader, PdfFileWriter
import re
import csv
from os import listdir
import os

## GLOBAL VARIABLES ##
inputs_dir = 'inputs'
contains_errors = False
contains_unknown = False
contains_pickups = False

csv_input_filenames = []
pdf_input_filenames = []

errors = []
actions = []

order_drivers = {} # key: order_id =, val: filename + driver
driver_export_filenames = [] # filename + driver
pdf_export_files = {}
pdf_input_files = {}
order_pages = {}
## GLOBAL VARIABLES END ##

# only appear in first page of Order
required_keywords = ["Order", "Date", "Shipping Method", "Tags", "Bill To"]
pickup_keywords = ["pickup", "pick up", "pick-up"] 

def process_deliveries_input(input_filename):
    driver_index = -1
    id_index = -1
    print("processing  " + input_filename)

    with open(inputs_dir + '/' + input_filename) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        
        for row in csv_reader:
            line_count += 1

            if line_count == 1: # headers row
                id_index, driver_index = process_header_row(row)
                continue

            order_id = row[id_index]
            if not row[id_index]:
                continue
            
            if row[driver_index] not in driver_export_filenames:
                driver_export_filenames.append(input_filename + "__" + row[driver_index])
            order_drivers[order_id] = input_filename + "__" + row[driver_index]

    print(" done")

def process_csv_inputs():
    for input_filename in csv_input_filenames:
        process_deliveries_input(input_filename)

def process_header_row(header_row):
    col_index = 0
    for header in header_row:
        if header == "Order ID" in header:
            id_index = col_index
        elif header == "Driver":
            driver_index = col_index
        col_index += 1
    if id_index == -1 or driver_index == -1:
        print("ERROR: Did not find target columns in csv")
        exit()
    return id_index, driver_index

def create_pdf_exports():
    for driver_filename in driver_export_filenames:
        pdf_export_files[driver_filename] = PdfFileWriter()

def create_pickups_pdf_export():
    pdf_export_files['pickups'] = PdfFileWriter()
    global contains_pickups
    contains_pickups = True

def create_error_pdf_export():
    pdf_export_files['ERRORS'] = PdfFileWriter()
    global contains_errors
    contains_errors = True

def create_unknown_pdf_export():
    pdf_export_files['UNKNOWN'] = PdfFileWriter()
    global contains_unknown
    contains_unknown = True

def open_pdf_inputs():
    for pdf_filename in pdf_input_filenames:
        pdf_input_files[pdf_filename] = open((inputs_dir + '/' + pdf_filename), 'rb')

def close_pdf_inputs():
    for pdf_file in pdf_input_files.values():
        pdf_file.close()

def process_pdf_inputs():
    for pdf_filename in pdf_input_files:
        print("processing  " + pdf_filename)
        process_pdf_input(pdf_input_files[pdf_filename], pdf_filename)

def process_pdf_input(pdf_file, filename):
    current_order_id = "PLACEHOLDER"
    current_driver = "UNKNOWN"

    pdfReader = PyPDF2.PdfFileReader(pdf_file)
    print(" No. Of Pages :", pdfReader.numPages)

    for page_index in range(pdfReader.numPages):
        page = pdfReader.getPage(page_index)
        text = page.extractText().replace('\n','')
        page_num = page_index +1

        # Check to see if the page contains all the keywords found on first page of orders
        contains_keywords = all(x in text for x in required_keywords)

        # Get text between 'Order ' and 'Date' and remove newlines
        order_id = text[text.find("Order ")+len("Order "):text.rfind("Date")].replace('\n','')
        
        # Check to see if order id looks valid
        id_looks_valid = len(order_id) < 12 and "#" in order_id

        driver = 'UNKNOWN'
        action = 'Added successfully'

        # # First page
        if contains_keywords:
            
            # order id is valid
            if order_id and id_looks_valid:
                current_order_id = order_id
                #  order in deliveries csv
                if order_id in order_drivers:
                    driver = order_drivers[order_id]
                # order NOT in deliveries csv
                else:
                    shipping_method = text[text.find("Shipping Method") +
                        len("Shipping Method"):text.rfind("Total Items")].replace('\n','')
                    # order is a PICKUP
                    if shipping_method and any(x in shipping_method.lower() for x in pickup_keywords):
                        driver = "pickups"
                        action = 'Not in deliveries and Shipping text found: ' + shipping_method
                        if not contains_pickups:
                            create_pickups_pdf_export()
                    # order is UNKNOWN
                    else:
                        driver = 'UNKNOWN'
                        action = 'CAUTION: Order ID not found in delivery input files'
                        if not contains_unknown:
                            create_unknown_pdf_export()
            # order id is invalid - ERRORED page
            else:
                errors.append('Unable to find Order ID on ' +
                    str(page_num) + ' in file: ' + filename)
                driver = 'ERROR'
                current_order_id = 'ERROR'
                action = 'ERROR: Unable to find Order ID from the first page of order'
                errors.append([filename, str(page_num), current_order_id, driver + '.pdf', action])
                if not contains_errors:
                    create_error_pdf_export()
            current_driver = driver
        else:
            driver = current_driver
            action = 'Not first page of Order. Added to same as previous page.'

        pdf_export_files[driver].addPage(page)
        actions.append([filename, str(page_num), current_order_id, driver + '.pdf', action])

    print(" done")

def close_pdf_exports():
    for driver in pdf_export_files:
        with open(driver + '.pdf', 'wb') as outfile:
            pdf_export_files[driver].write(outfile)

def find_inputs_from_subdir(suffix, path_to_dir=inputs_dir):
    filenames = listdir(path_to_dir)
    return [ filename for filename in filenames if filename.endswith( suffix ) ]

def scan_for_inputs():
    global csv_input_filenames
    global pdf_input_filenames
    csv_input_filenames = find_inputs_from_subdir(".csv")
    pdf_input_filenames = find_inputs_from_subdir(".pdf")
    csv_input_filenames.reverse()
    pdf_input_filenames.reverse()

    print(str(len(csv_input_filenames)) + " csv input files found: " + str(csv_input_filenames))
    print(str(len(pdf_input_filenames)) + " pdf input files found: " + str(pdf_input_filenames))

    if not csv_input_filenames:
        print("ERROR: No csv input files found!!")
        exit()
    if not pdf_input_filenames:
        print("ERROR: No pdf input files found!!")
        exit()

def create_reports():

    print('Creation Action Report')
    with open('_action_report.csv', 'w') as file:
        filewriter = csv.writer(file)
        filewriter.writerow(['Input PDF', 'Page num', 'Order ID', 'Export PDF', 'Action'])
        for row in actions:
            filewriter.writerow(row)
            # print(row)
    print(' done')

    if errors:
        print('Creation ERROR Report')
        with open('_error_report.csv', 'w') as file:
            filewriter = csv.writer(file)
            filewriter.writerow(['Input PDF', 'Page num', 'Order ID', 'Export PDF', 'Action'])
            for row in errors:
                filewriter.writerow(row)
                # print(row)
        print(' done')
    else:
        print('No errors')

def main():
    print(os.getcwd())
    # Search inputs folder for pdf and csv files
    scan_for_inputs()

    # Read and process csv files
    process_csv_inputs()

    # Open pdf input files
    open_pdf_inputs()

    # Create pdf files for exports
    create_pdf_exports()

    # Read pdf input and copy orders into correct export
    process_pdf_inputs()

    # Close off pdf exports
    close_pdf_exports()

    # Close off pdf inputs
    # Must be done after closing off pdf exports
    close_pdf_inputs()

    create_reports()

if __name__ == "__main__":
    # execute only if run as a script
    main()