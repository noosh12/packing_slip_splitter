import PyPDF2
from PyPDF2 import PdfFileReader, PdfFileWriter
import re
import csv
from os import listdir

## GLOBAL VARIABLES ##
inputs_dir = 'inputs'
pickups_pdf_filename = 'pickups'
error_pdf_filename = 'errors'
contains_errors = False

csv_input_filenames = []
pdf_input_filenames = []

order_drivers = {} # key: order_id =, val: filename + driver
driver_export_filenames = [] # filename + driver
pdf_export_files = {}
pdf_input_files = {}
order_pages = {}
## GLOBAL VARIABLES END ##

# only appear in first page of Order
required_keywords = ["Order", "Date", "Shipping Method", "Tags", "Bill To"] 

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
    # print(len(deliveries))
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
    # continue
    return id_index, driver_index

def create_pdf_exports():
    for driver_filename in driver_export_filenames:
        pdf_export_files[driver_filename] = PdfFileWriter()
    
    pdf_export_files['pickup'] = PdfFileWriter()

def create_error_pdf_export():
    pdf_export_files['error'] = PdfFileWriter()
    contains_errors = True
    # TODO create unknown pdf

def open_pdf_inputs():
    for pdf_filename in pdf_input_filenames:
        pdf_input_files[pdf_filename] = open((inputs_dir + '/' + pdf_filename), 'rb')

def close_pdf_inputs():
    for pdf_file in pdf_input_files.values():
        pdf_file.close()

def process_pdf_inputs():
    for pdf_filename in pdf_input_files:
        print("processing  " + pdf_filename)
        process_pdf_input(pdf_input_files[pdf_filename])

def process_pdf_input(pdf_file):

    current_order_id = "PLACEHOLDER"
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

        # TODO redo code below

        # First page of order
        if order_id and id_looks_valid and contains_keywords:
            order_pages[order_id] = str(page_num)
            # print("ADDED " + order_pages[order_id])
            current_order_id = order_id
            if order_id in order_drivers:
                driver = order_drivers[order_id]
                pdf_export_files[driver].addPage(page)
            continue

        # Not first page
        order_pages[current_order_id] = order_pages[current_order_id] + "," + str(page_num)
        if order_id in order_drivers:
            driver = order_drivers[order_id]
            # print(driver)
            pdf_export_files[driver].addPage(page)
        # else:
        #     print('WARNING: Found valid id in input pdf not in deliveries csv')
        #     print(order_id)

        # Keywords found but unable to find order id
        if contains_keywords:
            print('ERROR: Unable to find Order ID on page ' + page_num)

    print(" done")
    # for order in order_pages:
    #     print(order + " - " + order_pages[order])

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

    print(str(len(csv_input_filenames)) + " csv input files found: " + str(csv_input_filenames))
    print(str(len(pdf_input_filenames)) + " pdf input files found: " + str(pdf_input_filenames))

    if not csv_input_filenames:
        print("ERROR: No csv input files found!!")
        exit()
    if not pdf_input_filenames:
        print("ERROR: No pdf input files found!!")
        exit()

def main():
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

if __name__ == "__main__":
    # execute only if run as a script
    main()