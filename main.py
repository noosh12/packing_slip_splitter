import PyPDF2
from PyPDF2 import PdfFileReader, PdfFileWriter
import re
import csv
from os import listdir

monday_input = 'routes__1907__Canberra, New South Wales.csv'
pdf_input = 'OrderlyPrint Part 1.pdf'

inputs_dir = 'inputs'
csv_input_files = []
pdf_input_files = []

deliveries = {}
drivers = []
exports = {}
pdf_inputs = {}
order_pages = {}
# order_regex = "Order #(.*)"

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
            
            if row[driver_index] not in drivers:
                drivers.append(input_filename + "__" + row[driver_index])
            deliveries[order_id] = input_filename + "__" + row[driver_index]
    # print(len(deliveries))
    print(" done")

def process_deliveries_inputs():
    
    for input_file in csv_input_files:
        process_deliveries_input(input_file)

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
    for driver in drivers:
        exports[driver] = PdfFileWriter()
    # # print(drivers)
    # print(len(drivers))
    # print(len(exports))
    # for key in exports:
    #     print(key)

def open_pdf_inputs():
    for pdf in pdf_input_files:
        pdf_inputs[pdf] = open((inputs_dir + '/' + pdf), 'rb')
        # print(pdf)

    # pdf_inputs[pdf_input] = open(pdf_input, 'rb')

def close_pdf_inputs():
    # pdf_inputs[pdf_input].close()
    for pdf in pdf_inputs.values():
        pdf.close()

def process_pdf_inputs():
    for pdf in pdf_inputs:
        print("processing  " + pdf)
        process_pdf_input(pdf_inputs[pdf])

def process_pdf_input(pdf_file):

    current_order_id = "PLACEHOLDER"
    # pdfFileObject = open(pdf_input, 'rb')
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

        # todo FIX BELOW!!!!!!!!!!!!!!!!!!|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

        # First page of order
        if order_id and id_looks_valid and contains_keywords:
            order_pages[order_id] = str(page_num)
            # print("ADDED " + order_pages[order_id])
            current_order_id = order_id
            if order_id in deliveries:
                driver = deliveries[order_id]
                exports[driver].addPage(page)
            continue

        # Not first page
        order_pages[current_order_id] = order_pages[current_order_id] + "," + str(page_num)
        if order_id in deliveries:
            driver = deliveries[order_id]
            # print(driver)
            exports[driver].addPage(page)
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
    for driver in exports:
        with open(driver + '.pdf', 'wb') as outfile:
            exports[driver].write(outfile)

def find_inputs_from_subdir(suffix, path_to_dir=inputs_dir):
    filenames = listdir(path_to_dir)
    return [ filename for filename in filenames if filename.endswith( suffix ) ]

def scan_for_inputs():
    global csv_input_files
    global pdf_input_files
    csv_input_files = find_inputs_from_subdir(".csv")
    pdf_input_files = find_inputs_from_subdir(".pdf")

    for name in csv_input_files:
        print("csv file: " + name)
    for name in pdf_input_files:
        print("pdf file: " + name)

def main():

    scan_for_inputs()
    
    process_deliveries_inputs()
    create_pdf_exports()

    open_pdf_inputs()
    # process_pdf_input(pdf_inputs[pdf_input])
    process_pdf_inputs()

    close_pdf_exports()
    close_pdf_inputs()

if __name__ == "__main__":
    # execute only if run as a script
    main()