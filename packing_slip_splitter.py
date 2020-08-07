import PyPDF2
import re
import csv
import os
import sys
import time
import fpdf

class Order:
    def __init__(self, order_id, input_file, driver, stop_no):
        self.order_id = order_id
        self.input_file = input_file
        self.driver = driver
        self.export_file = input_file[:-4] + "__ALIAS_" + driver
        self.stop_no = stop_no if int(stop_no) > 9 else '0' + stop_no

        global driver_data
        if self.export_file not in driver_data:
            driver_data[self.export_file] = Driver(self.export_file)
        driver_data[self.export_file].add_order_id(self.order_id)

    def get_driver_stamp(self):
        global driver_data
        alias = getattr(driver_data[self.export_file], 'alias')
        return alias + '-' + self.stop_no

class Driver:
    def __init__(self, driver):
        self.name = driver
        self.orders = []

        global driver_count
        self.alias = alias_options[driver_count]
        driver_count += 1

    def add_order_id(self, order_id):
        self.orders.append(order_id)
    
    def get_full_export_name(self):
        return self.name.replace('__ALIAS', '__'+self.alias)


## GLOBAL VARIABLES START ##
inputs_folder = 'inputs/'
exports_folder = 'exports/'
stamp_file = '__delete_me__.pdf'
alias_options = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
driver_count = 0
total_order_count = 0
execution_dir = ''
contains_errors = False
contains_unknown = False
contains_pickups = False

csv_input_filenames = []
pdf_input_filenames = []

errors = []
actions = []

order_data = {} # key: order_id =, val: Order object
driver_data = {} # key: export filename =, val: Driver object

driver_export_filenames = [] # filename + driver
pdf_export_files = {}
pdf_input_files = {}
order_pages = {}
## GLOBAL VARIABLES END ##

# only appear in first page of Order
required_keywords = ["Order", "Date", "Shipping Method", "Tags", "Bill To"]
pickup_keywords = ["pickup", "pick up", "pick-up"] 

def get_directory():
    print('Determining current execution method...')
    global execution_dir
    if getattr(sys, 'frozen', False):
        # we are running in a bundle
        execution_dir = '/'.join(sys.executable.split('/')[:-1]) + '/'
        print('  Done! Running as a single-click app (bundle)')
    else:
        # we are running in a normal Python environment
        execution_dir = os.getcwd() + '/'
        print('  Done! Running as a Python file')
    print('  Corrected directory is:\n  ' + execution_dir)

    print('Checking if inputs folder exists in Corrected directory...')
    if os.path.exists(execution_dir + inputs_folder):
        print('  Success!')
    else:
        print('  ERROR! There is no inputs folder in the Corrected directory.')
        print('  Please create a folder named \'' + inputs_folder[:-1] + '\'')
        print('  The folder containing the input files should be:')
        print('    ' + execution_dir + inputs_folder)
        print('  exiting...')
        exit()
    
    print('Checking if exports folder exists in Corrected directory...')
    if os.path.exists(execution_dir + exports_folder):
        print('  Success!')
    else:
        print('  ERROR! There is no exports folder in the Corrected directory.')
        print('  Please create a folder named \'' + exports_folder[:-1] + '\'')
        print('  The folder that will contain the exports files should be:')
        print('    ' + execution_dir + exports_folder)
        print('  exiting...')
        exit()

def scan_for_inputs():
    print('Searching for input files...')
    
    global csv_input_filenames
    global pdf_input_filenames
    csv_input_filenames = find_inputs_from_subdir(".csv")
    pdf_input_filenames = find_inputs_from_subdir(".pdf")
    
    if not csv_input_filenames:
        print("  ERROR: No csv input files found!!")
        print('  exiting...')
        exit()
    if not pdf_input_filenames:
        print("  ERROR: No pdf input files found!!")
        print('  exiting...')
        exit()

    print('  Done!')
    csv_input_filenames_temp.reverse()
    pdf_input_filenames.reverse()
    for filename in csv_input_filenames_temp:
        if 'sunday' in filename.lower():
            csv_input_filenames.insert(0,filename)
        else:
            csv_input_filenames.append(filename)
    time.sleep(0.25)

    print('  ' + str(len(csv_input_filenames)) + " csv input files found: ")
    for filename in csv_input_filenames:
        print("    " + filename)

    time.sleep(0.25)
    print('  ' + str(len(pdf_input_filenames)) + " pdf input files found: ")
    for filename in pdf_input_filenames:
        print("    " + filename)

def find_inputs_from_subdir(suffix):
    filenames = os.listdir(execution_dir+inputs_folder)
    return [ filename for filename in filenames if filename.endswith( suffix ) ]

def process_csv_inputs():
    for input_filename in csv_input_filenames:
        process_csv_input(input_filename)
        time.sleep(0.25)

def process_csv_input(input_filename):
    driver_index = -1
    id_index = -1
    stop_no_index = -1
    print("Processing:  " + input_filename)
    full_filepath = execution_dir + inputs_folder + input_filename
    with open(full_filepath, 'r', encoding='mac_roman', newline='') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        order_count = 0
        
        for row in csv_reader:
            line_count += 1

            if line_count == 1: # headers row
                id_index, driver_index, stop_no_index = process_header_row(row)
                continue

            order_id = row[id_index]
            driver = row[driver_index]
            stop_no = row[stop_no_index]

            if not row[id_index]:
                continue

            filename_temp = input_filename[:-4] + "__" + driver
            
            if filename_temp not in driver_export_filenames:
                driver_export_filenames.append(filename_temp)
            if order_id not in order_data:
                order_data[order_id] = Order(order_id, input_filename, driver, stop_no)
            else:
                print('  ERROR! ' + str(order_id) + ' already exists in file: ' + getattr(order_data[order_id], 'input_file'))
                print('  exiting...')
                exit()
            order_count += 1

    print('  Done!')
    print('  ' + str(order_count) + ' delivery orders added')
    print('  ' + str(len(order_data)) + ' Total')

def process_header_row(header_row):
    col_index = 0
    id_index = -1
    driver_index = -1
    stop_no_index = -1
    for header_val in header_row:
        if header_val == "Order ID":
            id_index = col_index
        elif header_val == "Driver":
            driver_index = col_index
        elif header_val == "Stop Number":
            stop_no_index = col_index
        col_index += 1
    if id_index == -1 or driver_index == -1 or stop_no_index == -1:
        print("ERROR: Did not find target columns in csv")
        exit()
    return id_index, driver_index, stop_no_index

def create_pdf_exports():
    for driver_filename in driver_data: # key only
        pdf_export_files[driver_filename] = PyPDF2.PdfFileWriter()

def create_pickups_pdf_export():
    pdf_export_files['pickups'] = PyPDF2.PdfFileWriter()
    global contains_pickups
    contains_pickups = True

def create_error_pdf_export():
    pdf_export_files['ERRORS'] = PyPDF2.PdfFileWriter()
    global contains_errors
    contains_errors = True

def create_unknown_pdf_export():
    pdf_export_files['UNKNOWN'] = PyPDF2.PdfFileWriter()
    global contains_unknown
    contains_unknown = True

def open_pdf_inputs():
    for pdf_filename in pdf_input_filenames:
        pdf_input_files[pdf_filename] = open((execution_dir + inputs_folder + pdf_filename), 'rb')

def close_pdf_inputs():
    for pdf_file in pdf_input_files.values():
        pdf_file.close()

def process_pdf_inputs():
    for pdf_filename in pdf_input_files:
        print("Processing:  " + pdf_filename)
        process_pdf_input(pdf_input_files[pdf_filename], pdf_filename)

def process_pdf_input(pdf_file, filename):
    current_order_id = "PLACEHOLDER"
    current_driver = "UNKNOWN"
    global total_order_count
    prev_total_order_count = total_order_count
    time.sleep(0.25)

    pdfReader = PyPDF2.PdfFileReader(pdf_file)
    print("  No. Of Pages :", pdfReader.numPages)
    n_bar = 25

    for page_index in range(pdfReader.numPages):
        # Progress bar update
        j = (page_index + 1) / pdfReader.numPages
        sys.stdout.write('\r')
        sys.stdout.write(f"  [{'=' * int(n_bar * j):{n_bar}s}] {int(100 * j)}%   Current Orders: {total_order_count+1}")
        sys.stdout.flush()

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
                total_order_count += 1
                current_order_id = order_id
                #  order in deliveries csv
                if order_id in order_data:
                    driver = getattr(order_data[order_id], 'export_file')
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

        # Add stamp
        if driver != 'ERROR' and driver != 'UNKNOWN' and not (driver == 'pickups' and contains_keywords):
            create_order_stamp(current_order_id, contains_keywords, driver == 'pickups')
            order_stamp = open(stamp_file,'rb')
            stamp_pdf = PyPDF2.PdfFileReader(order_stamp)
            stamp_page = stamp_pdf.getPage(0)
            page.mergePage(stamp_page)

        pdf_export_files[driver].addPage(page)
        actions.append([filename, str(page_num), current_order_id, driver + '.pdf', action])
    print("\n  New Orders Processed: " + str(total_order_count-prev_total_order_count))
    print("  Done!")

def create_order_stamp(order_id, first_page, pickup):
    x_offset = 72 if first_page else 175
    y_offset = 10 if first_page else 0
    stamp = order_id if pickup else order_data[order_id].get_driver_stamp()

    pdf = fpdf.FPDF()
    pdf.add_page()
    pdf.set_font('Arial', 'B', 11)
    pdf.set_xy(x_offset, y_offset)
    pdf.cell(40, 10, stamp)
    pdf.output(stamp_file, 'F')

def close_pdf_exports():
    print(str(len(pdf_export_files))+ ' Export pdfs to create.\n  Exporting... ')
    non_drivers = ['UNKNOWN', 'ERROR', 'pickups']

    for driver in pdf_export_files:
        filename = driver
        if filename not in non_drivers:
            filename = driver_data[driver].get_full_export_name()
        with open(execution_dir + exports_folder + filename + '.pdf', 'wb') as outfile:
            pdf_export_files[driver].write(outfile)
        print('    ' + filename + '.pdf')
    print('  Done!')

def create_reports():
    print('Creating Action Report...')
    with open(execution_dir + '_action_report.csv', 'w') as file:
        filewriter = csv.writer(file)
        filewriter.writerow(['Input PDF', 'Page num', 'Order ID', 'Export PDF', 'Action'])
        for row in actions:
            filewriter.writerow(row)
            # print(row)
    print('  Done!')

    if errors:
        print('Creating ERROR Report as errors found.')
        with open(execution_dir + '_error_report.csv', 'w') as file:
            filewriter = csv.writer(file)
            filewriter.writerow(['Input PDF', 'Page num', 'Order ID', 'Export PDF', 'Action'])
            for row in errors:
                filewriter.writerow(row)
                # print(row)
        print('  Done!')
    else:
        print('Completed with no errors... No Error report to create!')

def main():
    print('--------------------------------------------------------------')
    print('--------------- Starting packing_slip_splitter ---------------')
    print('--------------------------------------------------------------')
    time.sleep(1.5)

    # single click app or python file?
    get_directory()
    time.sleep(1.5)
    
    # Search inputs folder for pdf and csv files
    scan_for_inputs()
    time.sleep(1.0)

    # Read and process csv files
    process_csv_inputs()
    time.sleep(0.5)

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

    time.sleep(0.25)
    create_reports()
    time.sleep(0.25)
    
    print("Total Orders: " + str(total_order_count))
    time.sleep(0.25)

    print('--------------------------------------------------------------')
    print('--------------- packing_slip_splitter Complete ---------------')
    print('--------------------------------------------------------------')

if __name__ == "__main__":
    # execute only if run as a script
    main()