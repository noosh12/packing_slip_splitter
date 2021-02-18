import PyPDF2
import re
import csv
import os
import sys
import time
import fpdf
import resource
from string import ascii_letters, digits

class Order:
    def __init__(self, order_id, deliveries_file, driver, stop_num):
        self.order_id = order_id
        self.deliveries_file = deliveries_file
        self.driver = driver
        self.export_file = deliveries_file[:-4] + "_" + driver
        self.stop_no = stop_num if int(stop_num) > 9 else '0' + stop_num
        self.pages = []
        self.pdf_file = ''
        self.order_exists_in_pdf = False
        self.order_added_to_export = False
        self.export_pdf_file = ''
        self.source_pages = ''

        if driver != 'N/A':
            global driver_data
            driver_data[self.export_file].add_order_id(self.order_id)

    def get_driver_stamp(self):
        global driver_data
        alias = getattr(driver_data[self.export_file], 'alias')
        return alias + '-' + self.stop_no

    def set_pdf_file(self, pdf_file):
        self.pdf_file = pdf_file
        self.order_exists_in_pdf = True

    def get_pdf_file(self):
        return self.pdf_file
    
    def order_exists(self):
        return self.order_exists_in_pdf

class Driver:
    def __init__(self, driver, alias):
        self.name = driver
        self.orders = []
        self.alias = alias

    def add_order_id(self, order_id):
        self.orders.append(order_id)
    
    def get_full_export_name(self):
        return self.alias + "__" + self.name
    
    def get_orders(self):
        print(str(self.orders))


## GLOBAL VARIABLES START ##
inputs_folder = 'inputs/'
exports_folder = 'exports/'
stamp_file = '__delete_me__.pdf'
alias_options = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
driver_count = 0
total_orders_found = 0
total_orders_added = 0
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
pickup_orders = [] # surname__**__order_id

driver_export_filenames = [] # filename + driver
pdf_export_files = {}
pdf_input_files = {}
order_pages = {}

# only appear in first page of Order
required_keywords = ["Order", "Date", "Shipping Method", "Tags", "Bill To"]
pickup_keywords = ["pickup", "pick up", "pick-up"] 

## GLOBAL VARIABLES END ##


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
    
    # the soft limit imposed by the current configuration
    # the hard limit imposed by the operating system.
    soft, hard = resource.getrlimit(resource.RLIMIT_NOFILE)
    if soft < 8000:
        print('Open files limit is too low: ' + str(soft))
        print('  Setting to 8192..')
        resource.setrlimit(resource.RLIMIT_NOFILE, (8192, hard))
        print('  Done!')

def scan_for_inputs():
    print('Searching for input files...')
    
    global csv_input_filenames
    global pdf_input_filenames
    csv_input_filenames_temp = find_inputs_from_subdir(".csv")
    pdf_input_filenames = find_inputs_from_subdir(".pdf")
    
    if not csv_input_filenames_temp:
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
    global driver_data
    print("Processing:  " + input_filename)
    full_filepath = execution_dir + inputs_folder + input_filename
    with open(full_filepath, 'r', encoding='mac_roman', newline='') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        order_count = 0
        driver = "PLACEHOLDER"
        driver_full = "PLACEHOLDER"
        
        for row in csv_reader:
            line_count += 1

            if line_count == 1: # headers row
                id_index, driver_index, stop_no_index = process_header_row(row)
                continue

            order_id = row[id_index]
            stop_no = row[stop_no_index]

            if not row[id_index]:
                continue
            
            if driver != row[driver_index]:
                driver = row[driver_index]
                global driver_count
                alias = alias_options[driver_count]
                driver_count += 1
                driver_full = input_filename[:-4] + "_" + driver
                driver_data[driver_full] = Driver(driver_full, alias)

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

    create_coversheet("pickups")
    coversheet = open(stamp_file,'rb')
    coversheet_pdf = PyPDF2.PdfFileReader(coversheet)
    coversheet_page = coversheet_pdf.getPage(0)
    pdf_export_files["pickups"].addPage(coversheet_page)

def create_error_pdf_export():
    pdf_export_files['ERRORS'] = PyPDF2.PdfFileWriter()
    global contains_errors
    contains_errors = True

    create_coversheet("ERRORS")
    coversheet = open(stamp_file,'rb')
    coversheet_pdf = PyPDF2.PdfFileReader(coversheet)
    coversheet_page = coversheet_pdf.getPage(0)
    pdf_export_files["ERRORS"].addPage(coversheet_page)

def create_unknown_pdf_export():
    pdf_export_files['UNKNOWN'] = PyPDF2.PdfFileWriter()
    global contains_unknown
    contains_unknown = True

    create_coversheet("UNKNOWN")
    coversheet = open(stamp_file,'rb')
    coversheet_pdf = PyPDF2.PdfFileReader(coversheet)
    coversheet_page = coversheet_pdf.getPage(0)
    pdf_export_files["UNKNOWN"].addPage(coversheet_page)

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
    is_pickup = False
    global total_orders_found
    global total_orders_added
    prev_total_order_count = total_orders_found
    time.sleep(0.25)

    pdfReader = PyPDF2.PdfFileReader(pdf_file)
    print("  No. Of Pages :", pdfReader.numPages)
    n_bar = 25

    for page_index in range(pdfReader.numPages):
        # Progress bar update
        j = (page_index + 1) / pdfReader.numPages
        sys.stdout.write('\r')
        sys.stdout.write(f"  [{'=' * int(n_bar * j):{n_bar}s}] {int(100 * j)}%   Current Order Total: {total_orders_found+1}")
        sys.stdout.flush()

        page = pdfReader.getPage(page_index)
        rawText = page.extractText()
        text = rawText.replace('\n','')
        page_num = page_index +1
        

        # Check to see if the page contains all the keywords found on first page of orders
        contains_keywords = all(x in text for x in required_keywords)

        # Get text between 'Order ' and 'Date' and remove newlines
        order_id = text[text.find("Order ")+len("Order "):text.rfind("Date")].replace('\n','')
        
        # Check to see if order id looks valid
        id_looks_valid = len(order_id) < 12 and "#" in order_id

        # # First page
        if contains_keywords:
            is_pickup = False

            # order id is valid
            if order_id and id_looks_valid:
                total_orders_found += 1
                current_order_id = order_id

                #  order in deliveries csv
                if order_id in order_data:
                    driver = getattr(order_data[order_id], 'export_file')
                    order_data[order_id].set_pdf_file(filename)
                    order_data[order_id].pages.append(page)
                    order_data[order_id].source_pages = str(page_num)

                # order NOT in deliveries csv
                else:
                    order_data[order_id] = Order(order_id, 'N/A', 'N/A', '00')
                    order_data[order_id].set_pdf_file(filename)
                    order_data[order_id].source_pages = str(page_num)

                    shipping_method = text[text.find("Shipping Method")
                        + len("Shipping Method"):text.rfind("Total Items")].replace('\n','')
                    
                    # order is a PICKUP
                    if shipping_method and any(x in shipping_method.lower() for x in pickup_keywords):
                        is_pickup = True
                        order_data[order_id].driver = 'pickup'
                        order_data[order_id].deliveries_file = 'Pickup detected...'
                        # Attempt to get the Ship To person's surname
                        shipping_name = rawText[rawText.find("Ship To\n")+len("Ship To\n"):rawText.rfind("Shipping Method")]
                        if "\n" in shipping_name:
                            for line in shipping_name.split("\n"):
                                if any(i.isdigit() for i in line): # first line of address
                                    break
                                last_name = line.strip()
                            shipping_name = shipping_name.split("\n")[0].strip()

                        if len(last_name) > 0:
                            if ' ' in last_name:
                                last_name = last_name.split(" ")[-1]
                        else:
                            last_name = " "

                        global pickup_orders
                        pickup_orders.append(last_name.lower() + "__**__" + order_id)
                        order_data[order_id].pages.append(page)

                        if not contains_pickups:
                            create_pickups_pdf_export()

                    # order is UNKNOWN
                    else:
                        total_orders_added += 1
                        current_order_id = 'UNKNOWN'
                        if not contains_unknown:
                            create_unknown_pdf_export()
                        pdf_export_files['UNKNOWN'].addPage(page)
                        order_data[order_id].driver = 'UNKNOWN'
                        order_data[order_id].order_added_to_export = True
                        order_data[order_id].pdf_file = filename
                        order_data[order_id].export_pdf_file = 'UNKNOWN'
                        order_data[order_id].deliveries_file = 'Not Found...'
                        order_data[order_id].pages.append(page)
            
            # order id is invalid - ERRORED page
            else:
                current_order_id = 'ERROR'
                if not contains_errors:
                    create_error_pdf_export()
                pdf_export_files['ERROR'].addPage(page)
                errors.append('Unable to extract the Order ID from  page. Page ' + str(page_num)
                    + ' from ' + filename + '. Added to Errors.pdf.')
        
        # second page
        else:
            if current_order_id == 'UNKNOWN':
                pdf_export_files['UNKNOWN'].addPage(page)
                order_data[current_order_id].pages.append(page)
                order_data[current_order_id].source_pages += "," + str(page_num)
                actions.append([filename, str(page_num), current_order_id, driver + '.pdf', 'CAUTION: Additional page of previous'])
            elif current_order_id == 'ERROR':
                pdf_export_files['ERROR'].addPage(page)
                errors.append('An additional page of the above. Page ' + str(page_num)
                    + ' from ' + filename + '. Added to Errors.pdf.')
            elif is_pickup:
                order_data[current_order_id].pages.append(page)
                order_data[current_order_id].source_pages += "," + str(page_num)
            else:
                order_data[current_order_id].pages.append(page)
                order_data[current_order_id].source_pages += "," + str(page_num)

    print("\n  New Orders Processed: " + str(total_orders_found-prev_total_order_count))
    print("  Done!")

def process_pdf_outputs():
    n_bar = 10
    global total_orders_found
    global total_orders_added
    print("Building output pdf files...")

    for driver in driver_data:
        create_coversheet(driver_data[driver].get_full_export_name())
        coversheet = open(stamp_file,'rb')
        coversheet_pdf = PyPDF2.PdfFileReader(coversheet)
        coversheet_page = coversheet_pdf.getPage(0)
        pdf_export_files[driver].addPage(coversheet_page)

        order_count = 1
        for order in reversed(driver_data[driver].orders):
            is_first_page = True
            j = (order_count) / len(driver_data[driver].orders)
            sys.stdout.write('\r')
            sys.stdout.write(f"  {driver_data[driver].name}  [{'=' * int(n_bar * j):{n_bar}s}] {int(100 * j)}%  Orders Added: {order_count}")
            sys.stdout.flush()
            order_count += 1
            total_orders_added += 1
            
            for page in order_data[order].pages:
                create_order_stamp(order, is_first_page, False)
                order_stamp = open(stamp_file,'rb')
                stamp_pdf = PyPDF2.PdfFileReader(order_stamp)
                stamp_page = stamp_pdf.getPage(0)
                page.mergePage(stamp_page)
                pdf_export_files[driver].addPage(page)

                is_first_page = False
            order_data[order].order_added_to_export = True
            order_data[order].export_pdf_file = driver_data[driver].get_full_export_name()
        print()
    
    # pickups
    pickup_orders.sort()
    order_count = 1
    for key in pickup_orders:
        is_first_page = True
        total_orders_added += 1
        order = key.split("__**__")[1]

        for page in order_data[order].pages:
            if is_first_page:
                j = (order_count) / len(pickup_orders)
                sys.stdout.write('\r')
                sys.stdout.write(f"  pickups  [{'=' * int(n_bar * j):{n_bar}s}] {int(100 * j)}%  Orders Added: {order_count}")
                sys.stdout.flush()
                order_count += 1
            else:
                create_order_stamp(order, is_first_page, True)
                order_stamp = open(stamp_file,'rb')
                stamp_pdf = PyPDF2.PdfFileReader(order_stamp)
                stamp_page = stamp_pdf.getPage(0)
                page.mergePage(stamp_page)
            pdf_export_files['pickups'].addPage(page)
            is_first_page = False
        order_data[order].order_added_to_export = True
        order_data[order].export_pdf_file = 'pickups'
    print()
        

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

def create_coversheet(text):
    x_offset = 50 if len(text) > 10 else 75
    y_offset = 130

    pdf = fpdf.FPDF()
    pdf.add_page()
    pdf.set_font('Courier', 'B', 20)
    pdf.set_xy(x_offset, y_offset)
    pdf.cell(40, 10, text)
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
        filewriter.writerow(['Order ID', 'Driver', 'Deliveries CSV File', 'Input PDF File'
            , 'Source Pages', 'Page Count', 'Added to Export', 'Export PDF File'])

        for order in order_data.values():
            filewriter.writerow([
                order.order_id,
                order.driver,
                order.deliveries_file,
                order.pdf_file,
                order.source_pages,
                str(len(order.pages)),
                str(order.order_added_to_export),
                order.export_pdf_file + '.pdf'
            ])

            if not order.order_added_to_export:
                errors.append(order.order_id + ' was never added to an export. Found on page(s) ' 
                    + order.source_pages + ' in ' + order.pdf_file)
    print('  Done!')

    if errors:
        print('Creating ERROR Report as errors found.')
        with open(execution_dir + '_error_report.csv', 'w') as file:
            filewriter = csv.writer(file)
            for error_text in errors:
                filewriter.writerow(error_text)
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

    process_pdf_outputs()

    # Close off pdf exports
    close_pdf_exports()

    # Close off pdf inputs
    # Must be done after closing off pdf exports
    close_pdf_inputs()

    time.sleep(0.25)
    create_reports()
    time.sleep(0.25)
    
    print("Total Orders found in input pdfs: " + str(total_orders_found))
    print("Total Orders added to output pdfs: " + str(total_orders_added))
    if(total_orders_found != total_orders_added):
        print("ERROR: Order totals do not match!!!")
    time.sleep(0.25)

    print('--------------------------------------------------------------')
    print('--------------- packing_slip_splitter Complete ---------------')
    print('--------------------------------------------------------------')

if __name__ == "__main__":
    # execute only if run as a script
    main()