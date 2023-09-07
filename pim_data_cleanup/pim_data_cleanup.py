"""
================
PIM Data Cleanup
================

@Author:  Thomas Vaughn
@Version: 1.0.0
@Date:    9.6.2023

A program designed to automate data cleanup activities which must be performed through a web application's UI through a
browser.  It was written in Python and utilizes Selenium WebDriver to interact with a web-based enterprise Product
Information Management (PIM) system.

* FOR ADDITIONAL DETAILS, PLEASE CONSULT THE README.

====================
Implementation Notes
====================

GENERAL NOTE ABOUT CODE COMPLETENESS AND APPLICATION IN OTHER ENVIRONMENTS:

        This program was written to be used in an enterprise PIM system, so I have modified and/or removed some of the
        particular DOM element references and other system-specific information for security purposes.

        I am sharing this code in the hopes that the framework will be useful for someone who is looking for a similar
        automation solution.

        Anyone wishing to use a similar approach should keep in mind that the work of identifying the reference data
        to help WebDriver find the elements you are trying to work with will still need to be done, as well as
        analyzing the state changes and flow used by the web application you are interacting with.

NET CONTENT:

        Version 1.0.0 is designed to be run after the PIM Net Content Fix program has already reviewed/updated the
        products.  It still checks the net content but only to see if there are any blank values or values which the
        Net Content Fix program flagged as blanks with a '-1'.  For any products with a blank or -1 in Net Content,
        this program checks the Company Net Content to see if it has a valid value it can use to update the Net Content
        field.  This gives us an opportunity to further improve the data in the Net Content field.

Start Availability Date Time:

        This program is designed to read and update information from the main grid in the PIM system.  However, if a
        value in the Start Availability Date Time is invalid, the value will appear blank from the main grid even if it
        is not blank.  In order to identify truly invalid records as efficiently as possible, the program first checks
        for validity via the main grid.  Then for any items which appear invalid or blank from the main grid, it will
        double-check the value by opening up the edit attribute dialog.  Even invalid values show up there.  This
        provides the ability to access the actual value of the Start Availability Date Time attribute to truly and
        accurately remove validation errors with the smallest amount of impact on processing time as possible.

Custom Waits:

        The built-in selenium waits were not working as expected when attempting to wait until particular DOM elements
        had the value we were waiting for values in a particular attribute.  It seems that the way the PIM system has
        coded certain elements makes it difficult to access them through standard routes in selenium.  As a workaround,
        this program includes custom waits in functions like check_lui_maingrid, get_row_id, and get_manufacturer_number
        which utilize try and except blocks to manage the program flow.  These custom waits are set to wait 1 decisecond
        between attempts.

Multi-Thread Program Flow Handling:

        The program uses the python keyboard module to assign a hotkey to stop the program and save the various files
        needed to ensure we have an accurate record of the program's activity.  This, however, means that the program
        utilizes multiple threads when a hotkey event is detected.  In order to manage the program flow between the
        threads, this program includes a threading.Event named freeze_event.  When the hotkey is detected, the secondary
        thread runs and finishes by setting the freeze_event.  The main thread is coded to alter several of its
        operations when freeze_event.is_set() including the main program loop as well as several blocks within the
        main loop.

lui_MainGrid:

        The element that primarily dictates program flow and made automation a challenge was the lui_MainGrid.  When
        loading/updating the contents of the grid, the style attribute goes...
            from:
                style="display: none"
            to:
                style="display: block"
            then once again to:
                style="display: none"

        This cycle from its resting state (none) to block to none again signals that the main grid is being reloaded.

            loaded = display: none
            loading = display: block

        Based selenium-waits for many activities in this program on this cycle since it was the most reliable method
        identified during development.

iframes:

        The program switches to the main grid iframe of the webpage before switching to the edit dialog iframe because
        removing this initial switch causes WebDriver to fail to find the edit dialog iframe
"""
import os
import threading

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
import sys
import time
import logging
import datetime
import ctypes
from ctypes import wintypes
import numpy
import keyboard

print("/////////////////////////////////////////////////////////////////////\n"
      "/                       PIM - Data Cleanup 1.0                      /\n"
      "/////////////////////////////////////////////////////////////////////\n"
      "\n"
      "This program automatically cycles through product records correcting the \n"
      "data in the following attributes:\n\n"
      "  * Manufacturer Number\n"
      "  * Start Availability Start Date\n"
      "  * Master GTIN\n"
      "  * Company Net Content\n"
      "  * Net Content\n\n"
      "=====================================================================\n"
      "\n")

input("Press ENTER to continue\n")

'''
Managing files and file directories 
'''
reviewed_path = "Reviewed Record Files"
corrected_path = "Corrected Record Files"
summary_path = "Activity Summary Files"
log_path = "Log Files"

# Check to see if each directory already exists
reviewed_exists = os.path.exists(reviewed_path)
corrected_exists = os.path.exists(corrected_path)
summary_exists = os.path.exists(summary_path)
log_exists = os.path.exists(log_path)

# If a directory doesn't exist yet, create the directory
if not reviewed_exists:
    os.makedirs(reviewed_path)
if not corrected_exists:
    os.makedirs(corrected_path)
if not summary_exists:
    os.makedirs(summary_path)
if not log_exists:
    os.makedirs(log_path)

# Global arrays to keep record of products reviewed and products updated
reviewed = ["Company Product Number", "Manufacturer Number", "Start Availability Date Time", "Master GTIN",
            "Net Content", "Company Net Content"]
fixed = ["Company Product Number", "Original Manufacturer Number", "Corrected Manufacturer",
         "Original Start Availability Date Time", "Corrected Start Availability Date Time", "Original Master GTIN",
         "Corrected Master GTIN", "Original Net Content", "Corrected Net Content", "Original Company Net Content",
         "Corrected Company Net Content"]

# Global counters
items_reviewed_counter = 0
items_fixed_counter = 0
errors_fixed_counter = 0

# Threading event to handle the save_and_quit function
freeze_event = threading.Event()


def init_logger():

    try:
        # Capture current datatime
        log_dtnow = datetime.datetime.now()

        # Format datetime
        log_date_time = log_dtnow.strftime("%m.%d.%Y_%H.%M.%S")

        # Logging setup
        logging.basicConfig(filename=log_path + '/pim_data_cleanup_' + log_date_time + '.log', level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %H:%M:%S::')

    except Exception as e:
        print("Exception occurred while attempting to initialize the logger.")
        print("Exiting program.")
        time.sleep(3)
        sys.exit()


"""
Flash Window
---
Utilizing FlashWindowEx to let user know when main program loop is completed  
"""
kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
user32 = ctypes.WinDLL('user32', use_last_error=True)


class WindowFlash(ctypes.Structure):

    _fields_ = (('csize', wintypes.UINT),
                ('hwnd', wintypes.HWND),
                ('flags', wintypes.DWORD),
                ('count', wintypes.UINT),
                ('timeout', wintypes.DWORD))

    def __init__(self, hwnd):
        self.csize = ctypes.sizeof(self)
        self.hwnd = hwnd
        self.flags = 0x00000002
        self.count = 9999999
        self.timeout = 0


def flash_window():

    try:

        h_wnd = kernel32.GetConsoleWindow()
        w_flash = WindowFlash(h_wnd)
        previous_state = user32.FlashWindowEx(ctypes.byref(w_flash))
        return previous_state

    except Exception as e:
        logging.error("Exception occurred during the flash_window method.\n", exc_info=True)
        print("Exception occurred during the flash_window method.")


def get_user_input_prerequisites():

    try:

        print("\nThe following actions must be performed before continuing:\n"
              "\n"
              " * Open Google Chrome in debug mode (batch script available)\n"
              " * Log in to the PIM system\n"
              " * Open the Staging Product repo\n"
              " * Make sure the Staging Product repo is the active tab in the PIM system\n"
              " * Switch the View Preference to 'Validation Automation'\n"
              " * Ensure the records per page is set to 50\n"
              "\n"
              "==========================================================================\n"              
              "\n")

        actions_confirmed = False  # loop variable

        # Prompt user to confirm that the prerequisite actions are completed
        while not actions_confirmed:
            actions_confirmed_input = input("Are the actions listed above complete? (y/n)")

            actions_confirmed_input = actions_confirmed_input.lower()

            if actions_confirmed_input == "y" or actions_confirmed_input == "yes":
                print("\nThank you.  Proceeding with program initialization.\n"
                      "\n"
                      "============================================================================\n"
                      "\n")

                actions_confirmed = True

            elif actions_confirmed_input == "n" or actions_confirmed_input == "no":
                print("\nPlease perform the following actions:\n"
                      "\n"
                      " * Open Google Chrome in debug mode (batch script available)\n"
                      " * Log in to the PIM system\n"
                      " * Open the Staging Product repo\n"
                      " * Make sure the Staging Product repo is the active tab in the PIM system\n"
                      " * Switch the View Preference to 'Validation Automation'\n"
                      " * Ensure the records per page is set to 50\n"
                      "\n"
                      "==========================================================================\n"              
                      "\n")
            else:
                print("\nInvalid input.  Please try again.\n")
                print("\nPlease perform the following actions:\n"
                      "\n"
                      " * Open Google Chrome in debug mode (batch script available)\n"
                      " * Log in to the PIM system\n"
                      " * Open the Staging Product repo\n"
                      " * Make sure the Staging Product repo is the active tab in the PIM system\n"
                      " * Switch the View Preference to 'Validation Automation'\n"
                      " * Ensure the records per page is set to 50\n"
                      "\n"
                      "==========================================================================\n"              
                      "\n")

    except Exception as e:
        print("Exception occurred during the get_user_input_prerequisites method.")
        print("Exiting program.")
        logging.error("Exception occurred during the get_user_input_prerequisites method.", exc_info=True)
        time.sleep(3)
        sys.exit()


def init_webdriver():

    try:

        """
        The following approach for initializing webdriver is not ideal, but it was used in lieu of the preferred 
        implementation in order to meet enterprise system requirements.
        """
        service = ChromeService()
        options = Options()
        options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        web_driver = webdriver.Chrome(service=service, options=options)

        return web_driver

    except:
        print("Error:: Exception occurred while attempting to initialize the Chrome WebDriver.")
        print("Exiting program.")
        logging.error("Exception occurred while attempting to initialize the Chrome WebDriver.", exc_info=True)
        time.sleep(3)
        sys.exit()


def parse_paging_info(driver):

    """
    Example of paging_info_elmt.text = "View 401 - 450 of 1,003,948"

    [substring_1: "View"]
    [substring_2: "401"]
    [substring_3: "-"]
    [substring_4: "450"]
    [substring_5: "of"]
    [substring_6: "1,003,948"]
    """

    try:
        # Retrieve the paging info element
        paging_info_elmt = driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/div")

        # Pull text from the paging_info_elmt
        paging_text = paging_info_elmt.text

        # Error handling
        if paging_text == "":
            print("Error:: Paging info was an empty string.")
            logging.error("Paging info was an empty string.")
            time.sleep(1)
            save_and_quit()
        elif paging_text is None:
            print("Error:: Paging info was None (null).")
            logging.error("Paging info was None (null).")
            time.sleep(1)
            save_and_quit()
        elif paging_text == 'No records to view':
            return paging_text
        else:
            # Parse the paging text
            paging_text_parsed = paging_text.split(' ')

            return paging_text_parsed

    except:
        print("Error:: Exception occurred while attempting to parse the paging info.")
        logging.error("Exception occurred while attempting to parse the paging info.", exc_info=True)
        time.sleep(1)
        save_and_quit()


def get_total_records(paging_info):

    try:
        # Extract total records and remove commas
        total_recs_text_raw = paging_info[5]  # index 5 = substring 6
        total_recs_text = total_recs_text_raw.replace(',', '')
        total_recs = int(total_recs_text)

        if total_recs < 1:
            logging.error("Error:: Parsed total record value less than 1.")
            print("Error:: Parsed total record value less than 1.")
            time.sleep(1)
            save_and_quit()

        return total_recs

    except:
        logging.error("Unable to extract total records from paging info.", exc_info=True)
        print("Error:: Unable to extract total records from paging info.")
        time.sleep(1)
        save_and_quit()


def get_total_pages(total_recs):

    try:
        # Calculate total pages
        if total_recs % 50 == 0:
            total_pgs = (int(total_recs) // 50)
        else:
            total_pgs = (int(total_recs) // 50) + 1

        return total_pgs

    except:
        logging.error("Unable to calculate total pages from paging info.", exc_info=True)
        print("Error:: Unable to calculate total pages from paging info.")
        time.sleep(1)
        save_and_quit()


def get_first_record_on_page(paging_info):

    try:
        # Extract current records and remove commas
        first_rec_text_raw = paging_info[1]  # index 1 = substring 2
        first_rec_text = first_rec_text_raw.replace(',', '')
        first_rec = int(first_rec_text)

        return first_rec

    except:
        logging.error("Unable to extract first record on page from paging info.", exc_info=True)
        print("Error:: Unable to extract first record on page from paging info.")
        time.sleep(1)
        save_and_quit()


def get_last_record_on_page(paging_info):

    try:
        # Extract current records and remove commas
        last_rec_text_raw = paging_info[3]  # index 3 = substring 4
        last_rec_text = last_rec_text_raw.replace(',', '')
        last_rec = int(last_rec_text)

        return last_rec

    except:
        logging.error("Unable to extract last record on page from paging info.", exc_info=True)
        print("Error:: Unable to extract last record on page from paging info.")
        time.sleep(1)
        save_and_quit()


def get_current_page(top_of_pg_rec):

    try:
        # Calculate starting page
        if top_of_pg_rec % 50 == 0:
            current_pg = (int(top_of_pg_rec) // 50)
        else:
            current_pg = (int(top_of_pg_rec) // 50) + 1

        return current_pg

    except:
        logging.error("Unable to calculate current page from paging info.", exc_info=True)
        print("Error:: Unable to calculate current page from paging info.")
        time.sleep(1)
        save_and_quit()


def print_initial_values(total_pgs, total_recs):

    # Print initial values used for main program loop
    print("\n\n======================\n"
          "\n"
          "Total Pages: " + str(total_pgs) + "\n"         
          "Total Records: " + str(total_recs) + "\n"
          "\n"
          "======================\n"
          "\n")

    # Print status message
    print("=========================================================================\n"
          "Initialization complete. Proceeding with product reviews and corrections.\n"
          "=========================================================================\n"
          "\n")
    # log initial values used for main program loop
    logging.info("\n\n======================\n"
          "Total Pages: " + str(total_pgs) + "\n"         
          "Total Records: " + str(total_recs) + "\n"
          "======================\n"
          "\n")
    # Log status message
    logging.info("\n\n===========================================================\n"
                 "Initialization complete. Proceeding with main program loop.\n"
                 "===========================================================\n"
                 "\n")


def check_lui_maingrid(web_driver):
    '''
    # Need to wait for the main grid to load completely before moving forward in main program loop.
    #  - Each time the main grid loads, the lui_MainGrid element goes from style="display: none" to
    #    style="display: block" and then back to style="display: none"
    #  - Attempts to use selenium waits and expected conditions like visibility_of_element_located
    #    failed to capture the actual loading status of the grid.
    '''

    # Variable to control program flow through the check_lui_maingraid method
    is_lui_finished = False

    # Setting loop variable to handle wait
    lui_maingrid_deciseconds = 0  # Loop variable

    while lui_maingrid_deciseconds < 600:  # 60 seconds
        try:
            # Locate the lui_maingrid element
            lui_element = web_driver.find_element(By.ID, "PLACEHOLDER")
            if "display: block" in lui_element.get_dom_attribute("style"):
                is_display_block = True
                lui_maingrid_deciseconds = 600
                continue
            else:
                lui_maingrid_deciseconds += 1
                time.sleep(.1)
        except:
            logging.error("lui main grid display never shifted to block\n")
            return False

    # Resetting loop variable
    lui_maingrid_deciseconds = 0  # Loop variable

    while lui_maingrid_deciseconds < 600:  # 60 seconds
        try:
            # Locate the lui_maingrid element
            lui_element = web_driver.find_element(By.ID, "PLACEHOLDER")
            if "display: none" in lui_element.get_dom_attribute("style"):
                is_display_none = True
                is_lui_finished = True
                lui_maingrid_deciseconds = 600
                continue
            else:
                lui_maingrid_deciseconds += 1
                time.sleep(.1)
        except:
            logging.error("lui main grid display never shifted to none\n")
            return False

    # If WebDriver registered expected lui_maingrid behavior
    if is_lui_finished:
        return True
    else:
        print("Error:: WebDriver was not able to verify that the maingrid loaded properly")
        logging.error("WebDriver was not able to verify that the maingrid loaded properly")
        return False


def check_lui_maingrid_click(web_driver):
    '''
    # Need to wait for the main grid to load completely before moving forward in main program loop.
    #  - Each time the main grid loads, the lui_MainGrid element goes from style="display: none" to
    #    style="display: block" and then back to style="display: none"
    #  - Specifically for use when preparing to check the checkbox for a record
    #  - Attempts to use selenium waits and expected conditions like visibility_of_element_located
    #    failed to capture the actual loading status of the grid.
    '''

    # Variable to control program flow through the check_lui_maingraid method
    is_lui_finished = False

    # Setting loop variable to handle wait
    lui_maingrid_deciseconds = 0  # Loop variable

    while lui_maingrid_deciseconds < 600:  # 60 seconds
        try:
            # Locate the lui_maingrid element
            lui_element = web_driver.find_element(By.ID, "PLACEHOLDER")
            if "display: none" in lui_element.get_dom_attribute("style"):
                is_display_none = True
                is_lui_finished = True
                lui_maingrid_deciseconds = 600
                continue
        except:
            lui_maingrid_deciseconds += 1
            time.sleep(.1)

    # If WebDriver registered expected lui_maingrid behavior
    if is_lui_finished:
        return True
    else:
        print("Error:: WebDriver was not able to verify that the maingrid loaded properly")
        logging.error("WebDriver was not able to verify that the maingrid loaded properly")
        return False


def get_row_id(web_driver, crnt_row):

    try:
        crnt_row = str(crnt_row)

        # Build an xpath for the top row element
        crnt_row_path = "PLACEHOLDER[" + crnt_row + "]"

        # Wait to make sure the driver finds the current table row
        try:
            crnt_row_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, crnt_row_path))
        except Exception as e:
            print("Error:: Current row element not found")
            logging.error("Current row element not found", exc_info=True)
            return "error"

        try:
            # Find the current row of the main grid table
            crnt_table_row = web_driver.find_element(By.XPATH, crnt_row_path)
        except Exception as e:
            print("Error:: WebDriver was not able to locate the current row of the main grid table")
            logging.error("WebDriver was not able to locate the current row of the main grid table", exc_info=True)
            return "error"

        # Setting loop variable to handle wait
        row_id_deciseconds = 0  # Loop variable
        crnt_row_id = ""

        while row_id_deciseconds < 200:  # 20 seconds
            try:
                # Pull the current row element's id
                crnt_row_id = crnt_table_row.get_attribute('id')
                row_id_deciseconds = 200
                continue
            except:
                row_id_deciseconds += 1
                time.sleep(.1)

        if crnt_row_id == "":
                logging.error("Unable to pull the current row element's id\n")
                print("Error:: Unable to pull the current row element's id\n")
                return "error"

    except Exception as e:
        print("Error:: Encountered problem with the get_row_id function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    return crnt_row_id


def get_Company_prod_number(web_driver, crnt_row_id):

    try:
        # Build an xpath for the Company product number field based on the current row id
        Company_prod_number_path = "//*[@id='" + crnt_row_id + "']/PLACEHOLDER"

        try:
            # Wait for the Company product number element in the current row to be found
            Company_prod_num_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, Company_prod_number_path))
        except Exception as e:
            print("Error:: Company prod number element not found")
            logging.error("Company prod number element not found", exc_info=True)
            return "error"

        try:
            # Find the Company product number element in the current row
            Company_prod_num_elmt = web_driver.find_element(By.XPATH, Company_prod_number_path)
        except Exception as e:
            print("Error:: WebDriver was not able to locate the Company prod number element of the current row")
            logging.error("Error:: WebDriver was not able to locate the Company prod number element of the current row", exc_info=True)
            return "error"

        # Setting loop variable to handle wait
        Company_prod_num_deciseconds = 0  # Loop variable
        Company_prod_num = "webdriver_failed_to_pull_Company_prod_number"

        while Company_prod_num_deciseconds < 200:  # 20 seconds
            try:
                # Get the Company prod number from the current row Company prod number field
                Company_prod_num = Company_prod_num_elmt.text
                Company_prod_num_deciseconds = 200
                continue
            except:
                Company_prod_num_deciseconds += 1
                time.sleep(.1)

        if Company_prod_num == "webdriver_failed_to_pull_Company_prod_number":
                print("Error:: Encountered problem with the get_Company_prod_number function.\n")
                logging.error("Unable to pull the current Company product number\n")
                return "error"

    except Exception as e:
        print("Error:: Encountered problem with the get_Company_prod_number function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    if Company_prod_num == "":
        Company_prod_num = "blank_in_main_grid"

    return Company_prod_num


def get_manufacturer_number(web_driver, crnt_row_id):

    try:

        # Build an xpath for the attribute field based on the current row id
        attribute_path = "//*[@id='" + crnt_row_id + "']/PLACEHOLDER"

        try:
            # Wait for the attribute element in the current row to be found
            attrib_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, attribute_path))
        except Exception as e:
            print("Error:: manufacturer number not found")
            logging.error("manufacturer number not found", exc_info=True)
            return "error"

        try:
            # Find the attribute element in the current row
            attribute_elmt = web_driver.find_element(By.XPATH, attribute_path)
        except Exception as e:
            print("Error:: WebDriver was not able to locate the manufacturer number element of the current row")
            logging.error("Error:: WebDriver was not able to locate the manufacturer number element of the current row",
                          exc_info=True)
            return "error"

        # Setting loop variable to handle wait
        attrib_deciseconds = 0  # Loop variable
        attrib_value = "webdriver_failed_to_pull_attribute_value"

        while attrib_deciseconds < 200:  # 20 seconds
            try:
                # Get the attribute value from the current row attribute field
                attrib_value = attribute_elmt.get_attribute('title')
                attrib_deciseconds = 200
                continue
            except:
                attrib_deciseconds += 1
                time.sleep(.1)

        if attrib_value == "webdriver_failed_to_pull_attribute_value":
                print("Error:: Encountered problem with the get_manufacturer_number function.\n")
                logging.error("Unable to pull the current manufacturer number\n")
                return "error"

    except Exception as e:
        print("Error:: Encountered problem with the get_manufacturer_number function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    if attrib_value == "":
        attrib_value = "blank_in_main_grid"

    return attrib_value


def get_brand_type(web_driver, crnt_row_id):

    try:

        # Build an xpath for the attribute field based on the current row id
        attribute_path = "//*[@id='" + crnt_row_id + "']/PLACEHOLDER"

        try:
            # Wait for the attribute element in the current row to be found
            attrib_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, attribute_path))
        except Exception as e:
            print("Error:: brand type not found")
            logging.error("brand type not found", exc_info=True)
            return "error"

        try:
            # Find the attribute element in the current row
            attribute_elmt = web_driver.find_element(By.XPATH, attribute_path)
        except Exception as e:
            print("Error:: WebDriver was not able to locate the brand type element of the current row")
            logging.error("Error:: WebDriver was not able to locate the brand type element of the current row",
                          exc_info=True)
            return "error"

        # Setting loop variable to handle wait
        attrib_deciseconds = 0  # Loop variable
        attrib_value = "webdriver_failed_to_pull_attribute_value"

        while attrib_deciseconds < 200:  # 20 seconds
            try:
                # Get the attribute value from the current row attribute field
                attrib_value = attribute_elmt.get_attribute('title')
                attrib_deciseconds = 200
                continue
            except:
                attrib_deciseconds += 1
                time.sleep(.1)

        if attrib_value == "webdriver_failed_to_pull_attribute_value":
                print("Error:: Encountered problem with the get_brand_type function.\n")
                logging.error("Unable to pull the current brand type\n")
                return "error"

    except Exception as e:
        print("Error:: Encountered problem with the get_brand_type function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    if attrib_value == "":
        attrib_value = "blank_in_main_grid"

    return attrib_value


def get_start_availability(web_driver, crnt_row_id):

    try:

        # Build an xpath for the attribute field based on the current row id
        attribute_path = "//*[@id='" + crnt_row_id + "']/PLACEHOLDER"

        try:
            # Wait for the attribute element in the current row to be found
            attrib_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, attribute_path))
        except Exception as e:
            print("Error:: start availability date time not found")
            logging.error("start availability date time not found", exc_info=True)
            return "error"

        try:
            # Find the attribute element in the current row
            attribute_elmt = web_driver.find_element(By.XPATH, attribute_path)
        except Exception as e:
            print("Error:: WebDriver was not able to locate the start availability date time element of the current row")
            logging.error("Error:: WebDriver was not able to locate the start availability date time element of the current row",
                          exc_info=True)
            return "error"

        # Setting loop variable to handle wait
        attrib_deciseconds = 0  # Loop variable
        attrib_value = "webdriver_failed_to_pull_attribute_value"

        while attrib_deciseconds < 200:  # 20 seconds
            try:
                # Get the attribute value from the current row attribute field
                attrib_value = attribute_elmt.get_attribute('title')
                attrib_deciseconds = 200
                continue
            except:
                attrib_deciseconds += 1
                time.sleep(.1)

        if attrib_value == "webdriver_failed_to_pull_attribute_value":
                print("Error:: Encountered problem with the get_start_availability function.\n")
                logging.error("Unable to pull the current start availability date time\n")
                return "error"

    except Exception as e:
        print("Error:: Encountered problem with the get_start_availability function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    if attrib_value == "":
        attrib_value = "blank_in_main_grid"

    return attrib_value


def get_master_gtin(web_driver, crnt_row_id):

    try:

        # Build an xpath for the attribute field based on the current row id
        attribute_path = "//*[@id='" + crnt_row_id + "']/PLACEHOLDER"

        try:
            # Wait for the attribute element in the current row to be found
            attrib_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, attribute_path))
        except Exception as e:
            print("Error:: master gtin not found")
            logging.error("master gtin not found", exc_info=True)
            return "error"

        try:
            # Find the attribute element in the current row
            attribute_elmt = web_driver.find_element(By.XPATH, attribute_path)
        except Exception as e:
            print("Error:: WebDriver was not able to locate the master gtin element of the current row")
            logging.error("Error:: WebDriver was not able to locate the master gtin element of the current row",
                          exc_info=True)
            return "error"

        # Setting loop variable to handle wait
        attrib_deciseconds = 0  # Loop variable
        attrib_value = "webdriver_failed_to_pull_attribute_value"

        while attrib_deciseconds < 200:  # 20 seconds
            try:
                # Get the attribute value from the current row attribute field
                attrib_value = attribute_elmt.get_attribute('title')
                attrib_deciseconds = 200
                continue
            except:
                attrib_deciseconds += 1
                time.sleep(.1)

        if attrib_value == "webdriver_failed_to_pull_attribute_value":
                print("Error:: Encountered problem with the get_master_gtin function.\n")
                logging.error("Unable to pull the current master gtin\n")
                return "error"

    except Exception as e:
        print("Error:: Encountered problem with the get_master_gtin function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    if attrib_value == "":
        attrib_value = "blank_in_main_grid"

    return attrib_value


def get_net_content(web_driver, crnt_row_id):

    try:

        # Build an xpath for the net content field based on the current row id
        net_content_path = "//*[@id='" + crnt_row_id + "']/PLACEHOLDER"

        try:
            # Wait for the net content element in the current row to be found
            net_cntnt_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, net_content_path))
        except Exception as e:
            print("Error:: net content not found")
            logging.error("net content not found", exc_info=True)
            return "error"

        try:
            # Find the net content element in the current row
            net_content_elmt = web_driver.find_element(By.XPATH, net_content_path)
        except Exception as e:
            print("Error:: WebDriver was not able to locate the net content element of the current row")
            logging.error("Error:: WebDriver was not able to locate the net content element of the current row", exc_info=True)
            return "error"

        # Setting loop variable to handle wait
        net_cntnt_deciseconds = 0  # Loop variable
        net_cntnt = "webdriver_failed_to_pull_net_cntnt"

        while net_cntnt_deciseconds < 200:  # 20 seconds
            try:
                # Get the net content value from the current row net content field
                net_cntnt = net_content_elmt.get_attribute('title')
                net_cntnt_deciseconds = 200
                continue
            except:
                net_cntnt_deciseconds += 1
                time.sleep(.1)

        if net_cntnt == "webdriver_failed_to_pull_net_cntnt":
                print("Error:: Encountered problem with the get_net_content function.\n")
                logging.error("Unable to pull the current net content\n")
                return "error"

    except Exception as e:
        print("Error:: Encountered problem with the get_net_content function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    if net_cntnt == "":
        net_cntnt = "blank_in_main_grid"

    return net_cntnt


def get_Company_net_content(web_driver, crnt_row_id):

    try:

        # Build an xpath for the attribute field based on the current row id
        attribute_path = "//*[@id='" + crnt_row_id + "']/PLACEHOLDER"

        try:
            # Wait for the attribute element in the current row to be found
            attrib_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, attribute_path))
        except Exception as e:
            print("Error:: Company net content not found")
            logging.error("Company net content not found", exc_info=True)
            return "error"

        try:
            # Find the attribute element in the current row
            attribute_elmt = web_driver.find_element(By.XPATH, attribute_path)
        except Exception as e:
            print("Error:: WebDriver was not able to locate the Company net content element of the current row")
            logging.error("Error:: WebDriver was not able to locate the Company net content element of the current row",
                          exc_info=True)
            return "error"

        # Setting loop variable to handle wait
        attrib_deciseconds = 0  # Loop variable
        attrib_value = "webdriver_failed_to_pull_attribute_value"

        while attrib_deciseconds < 200:  # 20 seconds
            try:
                # Get the attribute value from the current row attribute field
                attrib_value = attribute_elmt.get_attribute('title')
                attrib_deciseconds = 200
                continue
            except:
                attrib_deciseconds += 1
                time.sleep(.1)

        if attrib_value == "webdriver_failed_to_pull_attribute_value":
                print("Error:: Encountered problem with the get_Company_net_content function.\n")
                logging.error("Unable to pull the current Company net content\n")
                return "error"

    except Exception as e:
        print("Error:: Encountered problem with the get_Company_net_content function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    if attrib_value == "":
        attrib_value = "blank_in_main_grid"

    return attrib_value


def is_manufacturer_number_valid(attrib_value):
    """
    Based on the following requirements from the PIM Business Analyst:
    * Value should be numeric
    * Number of Digits = 6
    """

    try:
        # Variable for whether the value is valid, initialized to False
        is_valid_value = False

        # Check if the value contains only numeric digits
        try:
            int(attrib_value)
        except:
            return is_valid_value

        # Check if the field is blank
        if attrib_value == "":
            return is_valid_value

        elif attrib_value == "blank_in_main_grid":
            return is_valid_value

        # Check if manufacturer number has less than 6 characters
        elif len(attrib_value) < 6:
            return is_valid_value

        else:
            is_valid_value = True

    except Exception as e:
        print("Error:: Encountered problem with the is_manufacturer_number_valid function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    return is_valid_value


def is_start_availability_valid(attrib_value):
    """
    Based on the following requirements from the PIM Business Analyst:
    * The year of the datetime can't have 00 for the first two digits (ex: can't be 0007, must be 2007)
    * Based on product searches, oldest valid datetime has a year of 1982
    * Parse the datetime, if year less than 1982 then it's invalid
    * NOTE: Invalid values appear to be blank from the maingrid, so they should be considered invalid in this function
    """

    # Example of attribute value = "02/06/2018 00:00:01"

    try:

        # Variable for whether the value is valid, initialized to False
        is_valid_value = False

        # Check if attribute value is blank
        if attrib_value == "":
            return is_valid_value
        elif attrib_value == "blank_in_main_grid":
            return is_valid_value
        elif attrib_value is None:
            return is_valid_value

        # If not blank, parse the value and check for validity
        else:
            # Parse the attribute value by space, should yield a list with 2 elements
            attrib_value_parsed = attrib_value.split(' ')

        # pull the date from the parsed attribute value
        date = attrib_value_parsed[0]

        # check length of the date value
        if len(date) != 10:
            return is_valid_value
        else:
            # Parse the date by forward slash, should yield a list with 3 elements
            date_parsed = date.split('/')

        # pull the year from the parsed date
        year = date_parsed[2]

        try:
            # Convert year to an integer
            year = int(year)
        except:
            return is_valid_value

        # check if year is within valid range
        if year < 1982:
            return is_valid_value
        elif year >= 1982:
            is_valid_value = True

    except Exception as e:
        print("Error:: Encountered problem with the is_start_availability_valid function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    return is_valid_value


def doublecheck_start_availability(attrib_value):
    """
    Based on the following requirements from the PIM Business Analyst:
    * The year of the datetime can't have 00 for the first two digits (ex: can't be 0007, must be 2007)
    * Based on product searches, oldest valid datetime has a year of 1982
    * Parse the datetime, if year less than 1982 then it's invalid
    * This function is designed to check values taken from the edit attribute dialog (see main method for function call
      and arguments). Unlike values read from the maingrid, the edit attribute dialog properly displays values for the
      start availability date time even if they are invalid.  Therefore, blank values are considered valid in this
      function.
    """

    # Example of attribute value = "02/06/2018 00:00:01"

    try:

        # Variable for whether the value is valid, initialized to False
        is_valid_value = False

        # Check if attribute value is blank
        if attrib_value == "":
            is_valid_value = True
            return is_valid_value
        elif attrib_value == "blank_in_main_grid":
            is_valid_value = True
            return is_valid_value
        elif attrib_value is None:
            is_valid_value = True
            return is_valid_value

        # If not blank, parse the value and check for validity
        else:
            # Parse the attribute value by space, should yield a list with 2 elements
            attrib_value_parsed = attrib_value.split(' ')

        # pull the date from the parsed attribute value
        date = attrib_value_parsed[0]

        # check length of the date value
        if len(date) != 10:
            return is_valid_value
        else:
            # Parse the date by forward slash, should yield a list with 3 elements
            date_parsed = date.split('/')

        # pull the year from the parsed date
        year = date_parsed[2]

        try:
            # Convert year to an integer
            year = int(year)
        except:
            return is_valid_value

        # check if year is within valid range
        if year < 1982:
            return is_valid_value
        elif year >= 1982:
            is_valid_value = True

    except Exception as e:
        print("Error:: Encountered problem with the is_start_availability_valid function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    return is_valid_value


def is_master_gtin_valid(attrib_value):
    """
    Based on the following requirements from the PIM Business Analyst:
    * If value present, must be 14 digits long
    * If value present and less than 14 digits, pad the front with leading zeros until it is 14 digits
    * Blank values are allowed
    """

    try:
        # Variable for whether the value is valid, initialized to False
        is_valid_value = False

        # Check if the value is blank
        if attrib_value == "blank_in_main_grid":
            is_valid_value = True
            return is_valid_value

        # Check if the value contains only numeric digits
        try:
            int(attrib_value)
        except:
            return is_valid_value

        # Check if the value has valid character length
        if len(attrib_value) < 14:
            return is_valid_value
        elif len(attrib_value) > 14:
            return is_valid_value
        elif len(attrib_value) == 14:
            is_valid_value = True

    except Exception as e:
        print("Error:: Encountered problem with the is_master_gtin_valid function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    return is_valid_value


def is_net_content_blank(net_cntnt):
    """
    This function can be used instead of the is_net_content_valid function if the pim_net_content_fix program has
    already been run on the record set.  It will check for the following conditions:
    * net content is blank
    * net content is -1
    """

    try:
        # Variable for whether the value in net content is valid, initialized to False
        is_blank_nc = False

        # Check to see if net content is blank
        if net_cntnt == "":
            is_blank_nc = True
            return is_blank_nc
        elif net_cntnt == "blank_in_main_grid":
            is_blank_nc = True
            return is_blank_nc

        # Check to see if net content is -1 (means the net content was previously flagged as blank)
        elif net_cntnt == "-1":
            is_blank_nc = True
            return is_blank_nc

    except Exception as e:
        print("Error:: Encountered problem with the is_net_content_blank function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    return is_blank_nc


def is_Company_net_content_valid(attrib_value):
    """
    Based on the following requirements from the PIM Business Analyst:
    * Maximum Number of Characters = 9
    * No more than 2 digits to the right of the decimal
    * Should not be blank
    """

    try:
        # Variable for whether the value is valid, initialized to False
        is_valid_value = False

        # Check to see if the field is blank
        if attrib_value == "":
            return is_valid_value
        elif attrib_value == "blank_in_main_grid":
            return is_valid_value

        # Check to see if Company net content has more than 9 characters
        if len(attrib_value) > 9:
            return is_valid_value

        # Check to see if Company net content contains a dash
        has_dash = False

        if "-" in attrib_value:
            has_dash = True

        # Check to see if Company net content contains a decimal point
        has_decimal = False

        if "." in attrib_value:
            has_decimal = True

        # If the Company net content contains a dash, the only thing we need to check for is the character length
        if has_dash:

            if len(attrib_value) <= 9:
                is_valid_value = True
            else:
                is_valid_value = False

            # return now because the decimal checks and whole number checks don't matter if has_dash
            return is_valid_value

        if has_decimal:

            # Parse the Company net content
            Company_net_content_parsed = attrib_value.split('.')  # returns a list containing 2 strings

            # Get the digits to the left of the decimal point
            whole_number_digits = Company_net_content_parsed[0]  # use index to access the digits to the left of the decimal

            # Get the digits to the right of the decimal point
            decimal_digits = Company_net_content_parsed[1]  # use index to access the digits to the right of the decimal

            if len(attrib_value) <= 9 and len(decimal_digits) <= 2:
                is_valid_value = True

        else:  # no decimal or dash

            if len(attrib_value) <= 9:
                is_valid_value = True

    except Exception as e:
        print("Error:: Encountered problem with the is_Company_net_content_valid function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    return is_valid_value


def calculate_manufacturer_number(attrib_value):
    """
    Based on the following requirements from the PIM Business Analyst:
    * Value should be numeric
    * Number of Digits = 6
    * If less than 6 digits, pad the value with leading zeros
    """

    try:

        calculated_value = ""

        if attrib_value == "blank_in_main_grid":
            calculated_value = "000000"
            return calculated_value

        # check if value contains only numeric digits
        try:
            int(attrib_value)
        except:
            calculated_value = "not_a_number"

        # capture the length of the attribute value
        length = len(attrib_value)



        if length > 6:
            calculated_value = "exceeds character limit"

        if length < 6:
            calculated_value = attrib_value.zfill(6)

    except Exception as e:
        print("Error:: Encountered problem with the calculate_manufacturer_number function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    return calculated_value


def get_start_availability_dialog(web_driver, crnt_row_id):
    """
    Based on the following requirements from the PIM Business Analyst:
    * Maximum Number of Characters = 9
    * No more than 2 digits to the right of the decimal
    * Should not be blank
    ---------------------
    * All values with more than 2 decimal digits should be modified to include only 2
    * Blank values should be modified to -1 to show that the program has already checked the value
    """

    # Build an xpath for the attribute field based on the current row id
    attrib_path = "//*[@id='" + crnt_row_id + "']/PLACEHOLDER"

    try:
        # Find the attribute element in the current row
        attrib_elmt = web_driver.find_element(By.XPATH, attrib_path)
    except Exception as e:
        print("Error:: WebDriver was not able to locate the start availability date time element of the current row")
        logging.error("Error:: WebDriver was not able to locate the start availability date time element of the current row",
                      exc_info=True)
        return "error"

    # Click on the attribute element to open the edit attribute dialog
    if not click_start_availability(web_driver, crnt_row_id):
        print("Error:: WebDriver was not able to double-click the start availability date time element of the current row")
        logging.error("Error:: WebDriver was not able to double-click the start availability date time element of the current row",
                      exc_info=True)
        return "error"

    # Get the attribute value from the edit attribute dialog (only way to see values which exceed character limit)
    try:

        edit_attrib_path = "//div[PLACEHOLDER]"

        # Wait to make sure the driver finds the edit attribute dialog
        try:
            edit_attribute_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, edit_attrib_path))
        except Exception as e:
            print("Error:: Edit Attribute dialog not found")
            logging.error("Edit Attribute dialog not found", exc_info=True)
            return False

        # Setting loop variable to handle wait
        edit_attrib_deciseconds = 0  # Loop variable

        # Waiting until the edit attribute dialog display switches to block in the DOM
        while edit_attrib_deciseconds < 200:  # 20 seconds
            try:

                # Locate the edit attribute dialog element
                edit_attrib_element = web_driver.find_element(By.XPATH, edit_attrib_path)
                edit_elmt_display = edit_attrib_element.get_dom_attribute("style")

                if "display: block" in edit_elmt_display:
                    is_display_block = True
                    edit_attrib_deciseconds = 200
                    continue
                else:
                    edit_attrib_deciseconds += 1
                    time.sleep(.1)
            except:
                logging.error("edit attribute dialog display never shifted to block\n")
                return False

        # Resetting edit attribute seconds variable for use later in this method
        edit_attrib_deciseconds = 0

        # Wait to make sure the driver finds the edit attribute dialog
        try:
            edit_attribute_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Edit Attribute dialog not found")
            logging.error("Edit Attribute dialog not found", exc_info=True)
            return False

        try:
            # Find the edit attribute dialog
            edit_attribute_dialog = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the edit attribute dialog")
            logging.error("WebDriver was not able to locate the edit attribute dialog", exc_info=True)
            return False

        # Setting loop variable to handle wait
        edit_title_deciseconds = 0  # Loop variable

        # Waiting until the edit attribute title includes the text: Start Availability Date Time
        while edit_title_deciseconds < 200:  # 20 seconds
            try:
                is_net_content = "Start Availability Date Time" in edit_attribute_dialog.text

                if is_net_content:
                    edit_title_deciseconds = 200
                    continue
            except:
                edit_title_deciseconds += 1
                time.sleep(0.1)

        # Read the title of the edit attribute dialog to make sure we are about to edit the correct attribute
        edit_attribute_title = edit_attribute_dialog.text

        if edit_attribute_title != "Start Availability Date Time":
            print("Error:: The attribute selected was not Start Availability Date Time")
            logging.error("The attribute selected was not Start Availability Date Time")
            return False

        # Wait to make sure the driver finds the cell edit iframe
        try:
            cell_edit_iframe_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe"))
        except Exception as e:
            print("Error:: Cell Edit iframe not found")
            logging.error("Cell Edit iframe not found", exc_info=True)
            return False

        # Switch to the cell edit iframe
        iframe = web_driver.find_element(By.XPATH, "//*[@id='cellEditDiv']/iframe")
        web_driver.switch_to.frame(iframe)

        # Wait to make sure the driver finds the attribute value field
        try:
            attribute_value_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Attribute Value field not found")
            logging.error("Attribute Value field not found", exc_info=True)
            return False

        try:
            # Find the attribute value field
            attribute_value_field = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the attribute value field")
            logging.error("WebDriver was not able to locate the attribute value field", exc_info=True)
            return False

        try:
            # Retrieve the value from the attribute value field
            original_attribute_value = attribute_value_field.get_attribute('value')

            # Wait to make sure the driver finds the cancel button
            try:
                cancel_button_wait = WebDriverWait(web_driver, timeout=20).until(
                    lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
            except Exception as e:
                print("Error:: Cancel button not found")
                logging.error("Cancel button not found", exc_info=True)
                return False

            try:
                # Find the cancel button
                cancel_button = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
            except Exception as e:
                print("Error:: WebDriver was not able to locate the cancel button")
                logging.error("WebDriver was not able to locate the cancel button", exc_info=True)
                return False

            # Click the cancel button
            cancel_button.click()

            # Wait to make sure the driver finds the maingrid iframe
            try:
                maingrid_iframe_wait = WebDriverWait(web_driver, timeout=20).until(
                    lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe"))
            except Exception as e:
                print("Error:: maingrid iframe not found")
                logging.error("maingrid iframe not found", exc_info=True)
                return False

            # Switch to the main grid iframe
            iframe = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe")
            web_driver.switch_to.frame(iframe)

            # Used in loop below
            is_edit_attrib_closed = False

            # Waiting to make sure the edit attribute dialog is removed from the DOM before proceeding
            while edit_attrib_deciseconds < 200:  # 20 seconds
                # For this situation the except code block is the outcome we're looking for
                try:
                    # Check to see if web driver finds the element
                    edit_attrib_element = web_driver.find_element(By.XPATH, edit_attrib_path)
                    edit_attrib_deciseconds += 1
                    time.sleep(.1)
                except:
                    is_edit_attrib_closed = True
                    edit_attrib_deciseconds = 200
                    continue

            # If edit attribute dialog does not close
            if not is_edit_attrib_closed:
                print("Error:: Encountered problem with the get_start_availability_dialog function.\n")
                logging.error("The edit attribute dialog never left the DOM")
                return False

        except Exception as e:
            print("Error:: WebDriver was not able to retrieve the start availability date time value")
            logging.error("WebDriver was not able to retrieve the start availability date time value", exc_info=True)
            return False

    except Exception as e:
        print("Error:: Encountered problem with the get_start_availability_dialog function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    return original_attribute_value


def calculate_master_gtin(attrib_value):
    """
    Based on the following requirements from the PIM Business Analyst:
    * Value should be numeric
    * Number of Digits = 14
    * If less than 14 digits, pad the value with leading zeros
    """

    try:

        calculated_value = ""

        # check if value contains only numeric digits
        try:
            int(attrib_value)
        except:
            calculated_value = "not_a_number"

        # capture the length of the attribute value
        length = len(attrib_value)

        if length > 14:
            calculated_value = "exceeds character limit"

        if length < 14:
            calculated_value = attrib_value.zfill(14)

    except Exception as e:
        print("Error:: Encountered problem with the calculate_master_gtin function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"

    return calculated_value


def calculate_Company_net_content(web_driver, crnt_row_id):
    """
    Based on the following requirements from the PIM Business Analyst:
    * Maximum Number of Characters = 9
    * No more than 2 digits to the right of the decimal
    * Should not be blank
    ---------------------
    * All values with more than 2 decimal digits should be modified to include only 2
    * Blank values should be modified to -1 to show that the program has already checked the value
    """

    # Build an xpath for the attribute field based on the current row id
    attrib_path = "//*[@id='" + crnt_row_id + "']/PLACEHOLDER]"

    try:
        # Find the attribute element in the current row
        attrib_elmt = web_driver.find_element(By.XPATH, attrib_path)
    except Exception as e:
        print("Error:: WebDriver was not able to locate the Company net content element of the current row")
        logging.error("Error:: WebDriver was not able to locate the Company net content element of the current row",
                      exc_info=True)
        return "error"

    # Click on the attribute element to open the edit attribute dialog
    if not click_Company_net_content(web_driver, crnt_row_id):
        print("Error:: WebDriver was not able to double-click the Company net content element of the current row")
        logging.error("Error:: WebDriver was not able to double-click the Company net content element of the current row",
                      exc_info=True)
        return "error"

    # Get the attribute value from the edit attribute dialog (only way to see values which exceed character limit)
    try:

        edit_attrib_path = "//div[@aria-labelledby='PLACEHOLDER']"

        # Wait to make sure the driver finds the edit attribute dialog
        try:
            edit_attribute_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, edit_attrib_path))
        except Exception as e:
            print("Error:: Edit Attribute dialog not found")
            logging.error("Edit Attribute dialog not found", exc_info=True)
            return False

        # Setting loop variable to handle wait
        edit_attrib_deciseconds = 0  # Loop variable

        # Waiting until the edit attribute dialog display switches to block in the DOM
        while edit_attrib_deciseconds < 200:  # 20 seconds
            try:

                # Locate the edit attribute dialog element
                edit_attrib_element = web_driver.find_element(By.XPATH, edit_attrib_path)
                edit_elmt_display = edit_attrib_element.get_dom_attribute("style")

                if "display: block" in edit_elmt_display:
                    is_display_block = True
                    edit_attrib_deciseconds = 200
                    continue
                else:
                    edit_attrib_deciseconds += 1
                    time.sleep(.1)
            except:
                logging.error("edit attribute dialog display never shifted to block\n")
                return False

        # Resetting edit attribute seconds variable for use later in this method
        edit_attrib_deciseconds = 0

        # Wait to make sure the driver finds the edit attribute dialog
        try:
            edit_attribute_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Edit Attribute dialog not found")
            logging.error("Edit Attribute dialog not found", exc_info=True)
            return False

        try:
            # Find the edit attribute dialog
            edit_attribute_dialog = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the edit attribute dialog")
            logging.error("WebDriver was not able to locate the edit attribute dialog", exc_info=True)
            return False

        # Setting loop variable to handle wait
        edit_title_deciseconds = 0  # Loop variable

        # Waiting until the edit attribute title includes the text: Company Net Content
        while edit_title_deciseconds < 200:  # 20 seconds
            try:
                is_Company_net_content = "Company Net Content" in edit_attribute_dialog.text

                if is_Company_net_content:
                    edit_title_deciseconds = 200
                    continue
            except:
                edit_title_deciseconds += 1
                time.sleep(0.1)

        # Read the title of the edit attribute dialog to make sure we are about to edit the correct attribute
        edit_attribute_title = edit_attribute_dialog.text

        if edit_attribute_title != "Company Net Content":
            print("Error:: The attribute selected was not Company Net Content")
            logging.error("The attribute selected was not Company Net Content")
            return False

        # Wait to make sure the driver finds the cell edit iframe
        try:
            cell_edit_iframe_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe"))
        except Exception as e:
            print("Error:: Cell Edit iframe not found")
            logging.error("Cell Edit iframe not found", exc_info=True)
            return False

        # Switch to the cell edit iframe
        iframe = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe")
        web_driver.switch_to.frame(iframe)

        # Wait to make sure the driver finds the attribute value field
        try:
            attribute_value_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Attribute Value field not found")
            logging.error("Attribute Value field not found", exc_info=True)
            return False

        try:
            # Find the attribute value field
            attribute_value_field = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the attribute value field")
            logging.error("WebDriver was not able to locate the attribute value field", exc_info=True)
            return False

        try:
            # Retrieve the value from the attribute value field
            original_attribute_value = attribute_value_field.text

            # Wait to make sure the driver finds the cancel button
            try:
                cancel_button_wait = WebDriverWait(web_driver, timeout=20).until(
                    lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
            except Exception as e:
                print("Error:: Cancel button not found")
                logging.error("Cancel button not found", exc_info=True)
                return False

            try:
                # Find the cancel button
                cancel_button = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
            except Exception as e:
                print("Error:: WebDriver was not able to locate the cancel button")
                logging.error("WebDriver was not able to locate the cancel button", exc_info=True)
                return False

            # Click the cancel button
            cancel_button.click()

            # Wait to make sure the driver finds the maingrid iframe
            try:
                maingrid_iframe_wait = WebDriverWait(web_driver, timeout=20).until(
                    lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe"))
            except Exception as e:
                print("Error:: maingrid iframe not found")
                logging.error("maingrid iframe not found", exc_info=True)
                return False

            # Switch to the main grid iframe
            iframe = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe")
            web_driver.switch_to.frame(iframe)

            # Used in loop below
            is_edit_attrib_closed = False

            # Waiting to make sure the edit attribute dialog is removed from the DOM before proceeding
            while edit_attrib_deciseconds < 200:  # 20 seconds
                # For this situation the except code block is the outcome we're looking for
                try:
                    # Check to see if web driver finds the element
                    edit_attrib_element = web_driver.find_element(By.XPATH, edit_attrib_path)
                    edit_attrib_deciseconds += 1
                    time.sleep(.1)
                except:
                    is_edit_attrib_closed = True
                    edit_attrib_deciseconds = 200
                    continue

            # If edit attribute dialog does not close
            if not is_edit_attrib_closed:
                print("Error:: Encountered problem with the update_net_content function.\n")
                logging.error("The edit attribute dialog never left the DOM")
                return False

        except Exception as e:
            print("Error:: WebDriver was not able to retrieve the Company net content value")
            logging.error("WebDriver was not able to retrieve the Company net content value", exc_info=True)
            return False

        # Check to see if net content contains a dash
        has_dash = False
        if "-" in original_attribute_value:
            has_dash = True

        # Check to see if Company net content contains a decimal point
        has_decimal = False
        if "." in original_attribute_value:
            has_decimal = True

        if has_dash:

            # Parse the Company net content by the dash
            Company_net_content_parsed_dash = original_attribute_value.split('-')  # returns a list containing 2 strings

            # Get the value to the left of the dash
            min_value = Company_net_content_parsed_dash[0]  # use index to access the value to the left of the dash

            # Get the value to the right of the dash
            max_value = Company_net_content_parsed_dash[1]  # use index to access the value to the right of the dash

            # Check the character length of the minimum value
            if len(min_value) > 4:

                # use only the first 4 digits
                min_value = min_value[0:4]

                # if the 4th character in min_value is a decimal point, remove it
                if min_value[3] == ".":

                    # use only the first 3 digits
                    min_value = min_value[0:3]

            # Check the character length of the maximum value
            if len(min_value) > 4:

                # use only the 2 digits immediately to the right of decimal
                min_value = min_value[0:4]

                # if the 4th character in min_value is a decimal point, remove it
                if min_value[3] == ".":
                    # use only the first 3 digits
                    min_value = min_value[0:3]

            # build the fixed value
            fixed_attribute_value = min_value + "-" + max_value

            # go ahead and return the value since the additional checks / corrections are irrelevant when has_dash
            return fixed_attribute_value

        if has_decimal:

            # Parse the Company net content
            Company_net_content_parsed = original_attribute_value.split('.')  # returns a list containing 2 strings

            # Get the digits to the left of the decimal point
            whole_number_digits = Company_net_content_parsed[0]  # use index to access the digits to the left of the decimal

            # Get the digits to the right of the decimal point
            decimal_digits = Company_net_content_parsed[1]  # use index to access the digits to the right of the decimal

            if len(decimal_digits) > 2:

                # use only the 2 digits immediately to the right of decimal
                decimal_digits = decimal_digits[0:2]

                if len(whole_number_digits) > 6:
                    whole_number_digits = whole_number_digits[0:6]

                # build the fixed value
                fixed_attribute_value = whole_number_digits + "." + decimal_digits

                return fixed_attribute_value

        else:  # if no decimal (original value is blank)

            # update value to -1
            fixed_attribute_value = "-1"
            return fixed_attribute_value

    except Exception as e:
        print("Error:: Encountered problem with the calculate_Company_net_content function.\n")
        logging.error("Exception occurred", exc_info=True)
        return "error"


def click_manufacturer_number(web_driver, crnt_row_id):

    # Build an xpath for the attribute field based on the current row id
    attribute_path = "//*[@id='" + crnt_row_id + "']/PLACEHOLDER"

    # Wait to make sure the driver finds the attribute field for the current record
    try:
        attribute_wait = WebDriverWait(web_driver, timeout=20).until(
            lambda document: document.find_element(By.XPATH, attribute_path))
    except Exception as e:
        print("Error:: WebDriver was not able to locate the manufacturer number element of the current row")
        logging.error("WebDriver was not able to locate the manufacturer number element of the current row",
                      exc_info=True)
        return False

    try:
        # Find the attribute element in the current row
        attribute_elmt = web_driver.find_element(By.XPATH, attribute_path)
    except Exception as e:
        print("Error:: WebDriver was not able to locate the manufacturer number element of the current row")
        logging.error("WebDriver was not able to locate the manufacturer number element of the current row",
                      exc_info=True)
        return False

    # If frame is scrolled so attribute field is not visible, will scroll to the attribute element
    try:

        # Scroll to the element
        ActionChains(web_driver) \
            .scroll_to_element(attribute_elmt) \
            .perform()

    except Exception as e:
        print("Error:: could not scroll to manufacturer number field")
        logging.error("could not scroll to manufacturer number field", exc_info=True)
        return False

    # Make sure the main grid has finished loading
    if not check_lui_maingrid_click(web_driver):
        return False

    try:
        # Double-click the attribute field
        dbl_click_action = ActionChains(web_driver)
        dbl_click_action.double_click(attribute_elmt).perform()

    except Exception as e:
        print("Error:: WebDriver was not able to double-click the manufacturer number field")
        logging.error("WebDriver was not able to double-click the manufacturer number field", exc_info=True)
        return False

    # When function completes without any errors
    return True


def click_start_availability(web_driver, crnt_row_id):

    # Build an xpath for the attribute field based on the current row id
    attribute_path = "//*[@id='" + crnt_row_id + "']/PLACEHOLDER"

    # Wait to make sure the driver finds the attribute field for the current record
    try:
        attribute_wait = WebDriverWait(web_driver, timeout=20).until(
            lambda document: document.find_element(By.XPATH, attribute_path))
    except Exception as e:
        print("Error:: WebDriver was not able to locate the start availability element of the current row")
        logging.error("WebDriver was not able to locate the start availability element of the current row",
                      exc_info=True)
        return False

    try:
        # Find the attribute element in the current row
        attribute_elmt = web_driver.find_element(By.XPATH, attribute_path)
    except Exception as e:
        print("Error:: WebDriver was not able to locate the start availability element of the current row")
        logging.error("WebDriver was not able to locate the start availability element of the current row",
                      exc_info=True)
        return False

    # If frame is scrolled so attribute field is not visible, will scroll to the attribute element
    try:

        # Scroll to the element
        ActionChains(web_driver) \
            .scroll_to_element(attribute_elmt) \
            .perform()

    except Exception as e:
        print("Error:: could not scroll to start availability field")
        logging.error("could not scroll to start availability field", exc_info=True)
        return False

    # Make sure the main grid has finished loading
    if not check_lui_maingrid_click(web_driver):
        return False

    try:
        # Double-click the attribute field
        dbl_click_action = ActionChains(web_driver)
        dbl_click_action.double_click(attribute_elmt).perform()

    except Exception as e:
        print("Error:: WebDriver was not able to double-click the start availability field")
        logging.error("WebDriver was not able to double-click the start availability field", exc_info=True)
        return False

    # When function completes without any errors
    return True


def click_master_gtin(web_driver, crnt_row_id):

    # Build an xpath for the attribute field based on the current row id
    attribute_path = "//*[@id='" + crnt_row_id + "']/PLACEHOLDER"

    # Wait to make sure the driver finds the attribute field for the current record
    try:
        attribute_wait = WebDriverWait(web_driver, timeout=20).until(
            lambda document: document.find_element(By.XPATH, attribute_path))
    except Exception as e:
        print("Error:: WebDriver was not able to locate the master gtin element of the current row")
        logging.error("WebDriver was not able to locate the master gtin element of the current row",
                      exc_info=True)
        return False

    try:
        # Find the attribute element in the current row
        attribute_elmt = web_driver.find_element(By.XPATH, attribute_path)
    except Exception as e:
        print("Error:: WebDriver was not able to locate the master gtin element of the current row")
        logging.error("WebDriver was not able to locate the master gtin element of the current row",
                      exc_info=True)
        return False

    # If frame is scrolled so attribute field is not visible, will scroll to the attribute element
    try:

        # Scroll to the element
        ActionChains(web_driver) \
            .scroll_to_element(attribute_elmt) \
            .perform()

    except Exception as e:
        print("Error:: could not scroll to master gtin field")
        logging.error("could not scroll to master gtin field", exc_info=True)
        return False

    # Make sure the main grid has finished loading
    if not check_lui_maingrid_click(web_driver):
        return False

    try:
        # Double-click the attribute field
        dbl_click_action = ActionChains(web_driver)
        dbl_click_action.double_click(attribute_elmt).perform()

    except Exception as e:
        print("Error:: WebDriver was not able to double-click the master gtin field")
        logging.error("WebDriver was not able to double-click the master gtin field", exc_info=True)
        return False

    # When function completes without any errors
    return True


def click_Company_net_content(web_driver, crnt_row_id):

    # Build an xpath for the attribute field based on the current row id
    attribute_path = "//*[@id='" + crnt_row_id + "']/PLACEHOLDER"

    # Wait to make sure the driver finds the attribute field for the current record
    try:
        attribute_wait = WebDriverWait(web_driver, timeout=20).until(
            lambda document: document.find_element(By.XPATH, attribute_path))
    except Exception as e:
        print("Error:: WebDriver was not able to locate the Company net content element of the current row")
        logging.error("WebDriver was not able to locate the Company net content element of the current row",
                      exc_info=True)
        return False

    try:
        # Find the attribute element in the current row
        attribute_elmt = web_driver.find_element(By.XPATH, attribute_path)
    except Exception as e:
        print("Error:: WebDriver was not able to locate the Company net content element of the current row")
        logging.error("WebDriver was not able to locate the Company net content element of the current row",
                      exc_info=True)
        return False

    # If frame is scrolled so attribute field is not visible, will scroll to the attribute element
    try:

        # Scroll to the element
        ActionChains(web_driver) \
            .scroll_to_element(attribute_elmt) \
            .perform()

    except Exception as e:
        print("Error:: could not scroll to Company net content field")
        logging.error("could not scroll to Company net content field", exc_info=True)
        return False

    # Make sure the main grid has finished loading
    if not check_lui_maingrid_click(web_driver):
        return False

    try:
        # Double-click the attribute field
        dbl_click_action = ActionChains(web_driver)
        dbl_click_action.double_click(attribute_elmt).perform()

    except Exception as e:
        print("Error:: WebDriver was not able to double-click the Company net content field")
        logging.error("WebDriver was not able to double-click the Company net content field", exc_info=True)
        return False

    # When function completes without any errors
    return True


def click_net_content(web_driver, crnt_row_id):

    # Build an xpath for the net content field based on the current row id
    net_content_path = "//*[@id='" + crnt_row_id + "']/PLACEHOLDER"

    # Wait to make sure the driver finds the net content field for the current record
    try:
        net_content_wait = WebDriverWait(web_driver, timeout=20).until(
            lambda document: document.find_element(By.XPATH, net_content_path))
    except Exception as e:
        print("Error:: WebDriver was not able to locate the net content element of the current row")
        logging.error("WebDriver was not able to locate the net content element of the current row", exc_info=True)
        return False

    try:
        # Find the net content element in the current row
        net_content_elmt = web_driver.find_element(By.XPATH, net_content_path)
    except Exception as e:
        print("Error:: WebDriver was not able to locate the net content element of the current row")
        logging.error("WebDriver was not able to locate the net content element of the current row", exc_info=True)
        return False

    # If frame is scrolled so net content field is not visible, will scroll to the net content element
    try:

        # Scroll to the element
        ActionChains(web_driver) \
            .scroll_to_element(net_content_elmt) \
            .perform()

    except Exception as e:
        print("Error:: could not scroll to net content field")
        logging.error("could not scroll to net content field", exc_info=True)
        return False

    # Make sure the main grid has finished loading
    if not check_lui_maingrid_click(web_driver):
        return False

    try:
        # Double-click the net content field
        dbl_click_action = ActionChains(web_driver)
        dbl_click_action.double_click(net_content_elmt).perform()

    except Exception as e:
        print("Error:: WebDriver was not able to double-click the net content field")
        logging.error("WebDriver was not able to double-click the net content field", exc_info=True)
        return False

    # When function completes without any errors
    return True


def update_manufacturer_number(attrib_value, web_driver):
    """
    Based on the following requirements from the PIM Business Analyst:
    * Value should be numeric
    * Number of Digits = 6
    * If less than 6 digits, pad the value with leading zeros
    """

    try:

        edit_attrib_path = "//div[@aria-labelledby='PLACEHOLDER']"

        # Wait to make sure the driver finds the edit attribute dialog
        try:
            edit_attribute_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, edit_attrib_path))
        except Exception as e:
            print("Error:: Edit Attribute dialog not found")
            logging.error("Edit Attribute dialog not found", exc_info=True)
            return False

        # Setting loop variable to handle wait
        edit_attrib_deciseconds = 0  # Loop variable

        # Waiting until the edit attribute dialog display switches to block in the DOM
        while edit_attrib_deciseconds < 200:  # 20 seconds
            try:

                # Locate the edit attribute dialog element
                edit_attrib_element = web_driver.find_element(By.XPATH, edit_attrib_path)
                edit_elmt_display = edit_attrib_element.get_dom_attribute("style")

                if "display: block" in edit_elmt_display:
                    is_display_block = True
                    edit_attrib_deciseconds = 200
                    continue
                else:
                    edit_attrib_deciseconds += 1
                    time.sleep(.1)
            except:
                logging.error("edit attribute dialog display never shifted to block\n")
                return False

        # Resetting edit attribute seconds variable for use later in this method
        edit_attrib_deciseconds = 0

        # Wait to make sure the driver finds the edit attribute dialog
        try:
            edit_attribute_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Edit Attribute dialog not found")
            logging.error("Edit Attribute dialog not found", exc_info=True)
            return False

        try:
            # Find the edit attribute dialog
            edit_attribute_dialog = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the edit attribute dialog")
            logging.error("WebDriver was not able to locate the edit attribute dialog", exc_info=True)
            return False

        # Setting loop variable to handle wait
        edit_title_deciseconds = 0  # Loop variable

        # Waiting until the edit attribute title includes the text: Manufacturer Number
        while edit_title_deciseconds < 200:  # 20 seconds
            try:
                is_manufacturer_number = "Manufacturer Number" in edit_attribute_dialog.text

                if is_manufacturer_number:
                    edit_title_deciseconds = 200
                    continue
            except:
                edit_title_deciseconds += 1
                time.sleep(0.1)

        # Read the title of the edit attribute dialog to make sure we are about to edit the correct attribute
        edit_attribute_title = edit_attribute_dialog.text

        if edit_attribute_title != "Manufacturer Number":
            print("Error:: The attribute selected was not Manufacturer Number")
            logging.error("The attribute selected was not Manufacturer Number")
            return False

        # Wait to make sure the driver finds the cell edit iframe
        try:
            cell_edit_iframe_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe"))
        except Exception as e:
            print("Error:: Cell Edit iframe not found")
            logging.error("Cell Edit iframe not found", exc_info=True)
            return False

        # Switch to the cell edit iframe
        iframe = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe")
        web_driver.switch_to.frame(iframe)

        # Wait to make sure the driver finds the attribute value field
        try:
            attribute_value_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Attribute Value field not found")
            logging.error("Attribute Value field not found", exc_info=True)
            return False

        try:
            # Find the attribute value field
            attribute_value_field = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the attribute value field")
            logging.error("WebDriver was not able to locate the attribute value field", exc_info=True)
            return False

        try:
            # Clear the existing attribute value input text
            attribute_value_field.clear()

            # Enter the fixed attribute value
            attribute_value_field.send_keys(attrib_value)

        except:
            # Wait to make sure the driver finds the cancel button
            try:
                cancel_button_wait = WebDriverWait(web_driver, timeout=20).until(
                    lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
            except Exception as e:
                print("Error:: Cancel button not found")
                logging.error("Cancel button not found", exc_info=True)
                return False

            try:
                # Find the cancel button
                cancel_button = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
            except Exception as e:
                print("Error:: WebDriver was not able to locate the cancel button")
                logging.error("WebDriver was not able to locate the cancel button", exc_info=True)
                return False

            # Click the cancel button
            cancel_button.click()
            return False

        # Wait to make sure the driver finds the save button field
        try:
            save_button_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Save button not found")
            logging.error("Save button not found", exc_info=True)
            return False

        try:
            # Find the save button
            save_button = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the Save button")
            logging.error("WebDriver was not able to locate the Save button", exc_info=True)
            return False

        # Click the save button
        save_button.click()

        # Wait to make sure the driver finds the maingrid iframe
        try:
            maingrid_iframe_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe"))
        except Exception as e:
            print("Error:: maingrid iframe not found")
            logging.error("maingrid iframe not found", exc_info=True)
            return False

        # Switch to the main grid iframe
        iframe = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe")
        web_driver.switch_to.frame(iframe)

        # Used in loop below
        is_edit_attrib_closed = False

        # Waiting to make sure the edit attribute dialog is removed from the DOM before proceeding
        while edit_attrib_deciseconds < 200:  # 20 seconds
            # For this situation the except code block is the outcome we're looking for
            try:
                # Check to see if web driver finds the element
                edit_attrib_element = web_driver.find_element(By.XPATH, edit_attrib_path)
                edit_attrib_deciseconds += 1
                time.sleep(.1)
            except:
                is_edit_attrib_closed = True
                edit_attrib_deciseconds = 200
                continue

        # When function completes without any errors
        if is_edit_attrib_closed:
            return True
        else:
            print("Error:: Encountered problem with the update_manufacturer_number function.\n")
            logging.error("The edit attribute dialog never left the DOM")
            return False

    except Exception as e:
        print("Error:: Exception encountered when attempting to update the manufacturer number")
        logging.error("Exception encountered when attemtping to update the manufacturer number", exc_info=True)
        return False


def update_start_availability(web_driver):
    """
    Based on the following requirements from the PIM Business Analyst:
    * The year of the datetime can't have 00 for the first two digits (ex: can't be 0007, must be 2007)
    * Based on product searches, oldest valid datetime has a year of 1982
    * Parse the datetime, if year less than 1982, clear the datetime completely
    """

    try:

        edit_attrib_path = "//div[@aria-labelledby='PLACEHOLDER']"

        # Wait to make sure the driver finds the edit attribute dialog
        try:
            edit_attribute_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, edit_attrib_path))
        except Exception as e:
            print("Error:: Edit Attribute dialog not found")
            logging.error("Edit Attribute dialog not found", exc_info=True)
            return False

        # Setting loop variable to handle wait
        edit_attrib_deciseconds = 0  # Loop variable

        # Waiting until the edit attribute dialog display switches to block in the DOM
        while edit_attrib_deciseconds < 200:  # 20 seconds
            try:

                # Locate the edit attribute dialog element
                edit_attrib_element = web_driver.find_element(By.XPATH, edit_attrib_path)
                edit_elmt_display = edit_attrib_element.get_dom_attribute("style")

                if "display: block" in edit_elmt_display:
                    is_display_block = True
                    edit_attrib_deciseconds = 200
                    continue
                else:
                    edit_attrib_deciseconds += 1
                    time.sleep(.1)
            except:
                logging.error("edit attribute dialog display never shifted to block\n")
                return False

        # Resetting edit attribute seconds variable for use later in this method
        edit_attrib_deciseconds = 0

        # Wait to make sure the driver finds the edit attribute dialog
        try:
            edit_attribute_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Edit Attribute dialog not found")
            logging.error("Edit Attribute dialog not found", exc_info=True)
            return False

        try:
            # Find the edit attribute dialog
            edit_attribute_dialog = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the edit attribute dialog")
            logging.error("WebDriver was not able to locate the edit attribute dialog", exc_info=True)
            return False

        # Setting loop variable to handle wait
        edit_title_deciseconds = 0  # Loop variable

        # Waiting until the edit attribute title includes the text: Start Availability Date Time
        while edit_title_deciseconds < 200:  # 20 seconds
            try:
                is_start_availability = "Start Availability Date Time" in edit_attribute_dialog.text

                if is_start_availability:
                    edit_title_deciseconds = 200
                    continue
            except:
                edit_title_deciseconds += 1
                time.sleep(0.1)

        # Read the title of the edit attribute dialog to make sure we are about to edit the correct attribute
        edit_attribute_title = edit_attribute_dialog.text

        if edit_attribute_title != "Start Availability Date Time":
            print("Error:: The attribute selected was not Start Availability Date Time")
            logging.error("The attribute selected was not Start Availability Date Time")
            return False

        # Wait to make sure the driver finds the cell edit iframe
        try:
            cell_edit_iframe_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe"))
        except Exception as e:
            print("Error:: Cell Edit iframe not found")
            logging.error("Cell Edit iframe not found", exc_info=True)
            return False

        # Switch to the cell edit iframe
        iframe = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe")
        web_driver.switch_to.frame(iframe)

        # Wait to make sure the driver finds the attribute value field
        try:
            attribute_value_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Attribute Value field not found")
            logging.error("Attribute Value field not found", exc_info=True)
            return False

        try:
            # Find the attribute value field
            attribute_value_field = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the attribute value field")
            logging.error("WebDriver was not able to locate the attribute value field", exc_info=True)
            return False

        try:
            # Clear the existing attribute value input text
            attribute_value_field.clear()

        except:
            # Wait to make sure the driver finds the cancel button
            try:
                cancel_button_wait = WebDriverWait(web_driver, timeout=20).until(
                    lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
            except Exception as e:
                print("Error:: Cancel button not found")
                logging.error("Cancel button not found", exc_info=True)
                return False

            try:
                # Find the cancel button
                cancel_button = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
            except Exception as e:
                print("Error:: WebDriver was not able to locate the cancel button")
                logging.error("WebDriver was not able to locate the cancel button", exc_info=True)
                return False

            # Click the cancel button
            cancel_button.click()
            return False

        # Wait to make sure the driver finds the save button field
        try:
            save_button_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Save button not found")
            logging.error("Save button not found", exc_info=True)
            return False

        try:
            # Find the save button
            save_button = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the Save button")
            logging.error("WebDriver was not able to locate the Save button", exc_info=True)
            return False

        # Click the save button
        save_button.click()

        # Wait to make sure the driver finds the maingrid iframe
        try:
            maingrid_iframe_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe"))
        except Exception as e:
            print("Error:: maingrid iframe not found")
            logging.error("maingrid iframe not found", exc_info=True)
            return False

        # Switch to the main grid iframe
        iframe = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe")
        web_driver.switch_to.frame(iframe)

        # Used in loop below
        is_edit_attrib_closed = False

        # Waiting to make sure the edit attribute dialog is removed from the DOM before proceeding
        while edit_attrib_deciseconds < 200:  # 20 seconds
            # For this situation the except code block is the outcome we're looking for
            try:
                # Check to see if web driver finds the element
                edit_attrib_element = web_driver.find_element(By.XPATH, edit_attrib_path)
                edit_attrib_deciseconds += 1
                time.sleep(.1)
            except:
                is_edit_attrib_closed = True
                edit_attrib_deciseconds = 200
                continue

        # When function completes without any errors
        if is_edit_attrib_closed:
            return True
        else:
            print("Error:: Encountered problem with the update_start_availability function.\n")
            logging.error("The edit attribute dialog never left the DOM")
            return False

    except Exception as e:
        print("Error:: Exception encountered when attempting to update the start availability date time")
        logging.error("Exception encountered when attemtping to update the start availability date time", exc_info=True)
        return False


def update_master_gtin(attrib_value, web_driver):
    """
    Based on the following requirements from the PIM Business Analyst:
    * Value should be numeric
    * Number of Digits = 14
    * If less than 14 digits, pad the value with leading zeros
    """

    try:

        edit_attrib_path = "//div[@aria-labelledby='PLACEHOLDER']"

        # Wait to make sure the driver finds the edit attribute dialog
        try:
            edit_attribute_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, edit_attrib_path))
        except Exception as e:
            print("Error:: Edit Attribute dialog not found")
            logging.error("Edit Attribute dialog not found", exc_info=True)
            return False

        # Setting loop variable to handle wait
        edit_attrib_deciseconds = 0  # Loop variable

        # Waiting until the edit attribute dialog display switches to block in the DOM
        while edit_attrib_deciseconds < 200:  # 20 seconds
            try:

                # Locate the edit attribute dialog element
                edit_attrib_element = web_driver.find_element(By.XPATH, edit_attrib_path)
                edit_elmt_display = edit_attrib_element.get_dom_attribute("style")

                if "display: block" in edit_elmt_display:
                    is_display_block = True
                    edit_attrib_deciseconds = 200
                    continue
                else:
                    edit_attrib_deciseconds += 1
                    time.sleep(.1)
            except:
                logging.error("edit attribute dialog display never shifted to block\n")
                return False

        # Resetting edit attribute seconds variable for use later in this method
        edit_attrib_deciseconds = 0

        # Wait to make sure the driver finds the edit attribute dialog
        try:
            edit_attribute_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Edit Attribute dialog not found")
            logging.error("Edit Attribute dialog not found", exc_info=True)
            return False

        try:
            # Find the edit attribute dialog
            edit_attribute_dialog = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the edit attribute dialog")
            logging.error("WebDriver was not able to locate the edit attribute dialog", exc_info=True)
            return False

        # Setting loop variable to handle wait
        edit_title_deciseconds = 0  # Loop variable

        # Waiting until the edit attribute title includes the text: Master GTIN
        while edit_title_deciseconds < 200:  # 20 seconds
            try:
                is_master_gtin = "Master GTIN" in edit_attribute_dialog.text

                if is_master_gtin:
                    edit_title_deciseconds = 200
                    continue
            except:
                edit_title_deciseconds += 1
                time.sleep(0.1)

        # Read the title of the edit attribute dialog to make sure we are about to edit the correct attribute
        edit_attribute_title = edit_attribute_dialog.text

        if edit_attribute_title != "Master GTIN":
            print("Error:: The attribute selected was not Master GTIN")
            logging.error("The attribute selected was not Master GTIN")
            return False

        # Wait to make sure the driver finds the cell edit iframe
        try:
            cell_edit_iframe_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe"))
        except Exception as e:
            print("Error:: Cell Edit iframe not found")
            logging.error("Cell Edit iframe not found", exc_info=True)
            return False

        # Switch to the cell edit iframe
        iframe = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe")
        web_driver.switch_to.frame(iframe)

        # Wait to make sure the driver finds the attribute value field
        try:
            attribute_value_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Attribute Value field not found")
            logging.error("Attribute Value field not found", exc_info=True)
            return False

        try:
            # Find the attribute value field
            attribute_value_field = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the attribute value field")
            logging.error("WebDriver was not able to locate the attribute value field", exc_info=True)
            return False

        try:
            # Clear the existing attribute value input text
            attribute_value_field.clear()

            # Enter the fixed attribute value
            attribute_value_field.send_keys(attrib_value)

        except:
            # Wait to make sure the driver finds the cancel button
            try:
                cancel_button_wait = WebDriverWait(web_driver, timeout=20).until(
                    lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
            except Exception as e:
                print("Error:: Cancel button not found")
                logging.error("Cancel button not found", exc_info=True)
                return False

            try:
                # Find the cancel button
                cancel_button = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
            except Exception as e:
                print("Error:: WebDriver was not able to locate the cancel button")
                logging.error("WebDriver was not able to locate the cancel button", exc_info=True)
                return False

            # Click the cancel button
            cancel_button.click()
            return False

        # Wait to make sure the driver finds the save button field
        try:
            save_button_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Save button not found")
            logging.error("Save button not found", exc_info=True)
            return False

        try:
            # Find the save button
            save_button = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the Save button")
            logging.error("WebDriver was not able to locate the Save button", exc_info=True)
            return False

        # Click the save button
        save_button.click()

        # Wait to make sure the driver finds the maingrid iframe
        try:
            maingrid_iframe_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe"))
        except Exception as e:
            print("Error:: maingrid iframe not found")
            logging.error("maingrid iframe not found", exc_info=True)
            return False

        # Switch to the main grid iframe
        iframe = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe")
        web_driver.switch_to.frame(iframe)

        # Used in loop below
        is_edit_attrib_closed = False

        # Waiting to make sure the edit attribute dialog is removed from the DOM before proceeding
        while edit_attrib_deciseconds < 200:
            # For this situation the except code block is the outcome we're looking for
            try:
                # Check to see if web driver finds the element
                edit_attrib_element = web_driver.find_element(By.XPATH, edit_attrib_path)
                edit_attrib_deciseconds += 1
                time.sleep(.1)
            except:
                is_edit_attrib_closed = True
                edit_attrib_deciseconds = 200
                continue

        # When function completes without any errors
        if is_edit_attrib_closed:
            return True
        else:
            print("Error:: Encountered problem with the update_master_gtin function.\n")
            logging.error("The edit attribute dialog never left the DOM")
            return False

    except Exception as e:
        print("Error:: Exception encountered when attempting to update the master gtin")
        logging.error("Exception encountered when attemtping to update the master gtinr", exc_info=True)
        return False


def update_Company_net_content(attrib_value, web_driver):
    """
    Based on the following requirements from the PIM Business Analyst:
    * Maximum Number of Characters = 9
    * No more than 2 digits to the right of the decimal
    * Should not be blank
    ---------------------
    * All values with more than 2 decimal digits should be modified to include only 2
    * Blank values should be modified to -1 to show that the program has already checked the value
    """

    try:

        edit_attrib_path = "//div[@aria-labelledby='PLACEHOLDER']"

        # Wait to make sure the driver finds the edit attribute dialog
        try:
            edit_attribute_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, edit_attrib_path))
        except Exception as e:
            print("Error:: Edit Attribute dialog not found")
            logging.error("Edit Attribute dialog not found", exc_info=True)
            return False

        # Setting loop variable to handle wait
        edit_attrib_deciseconds = 0  # Loop variable

        # Waiting until the edit attribute dialog display switches to block in the DOM
        while edit_attrib_deciseconds < 200:  # 20 seconds
            try:

                # Locate the edit attribute dialog element
                edit_attrib_element = web_driver.find_element(By.XPATH, edit_attrib_path)
                edit_elmt_display = edit_attrib_element.get_dom_attribute("style")

                if "display: block" in edit_elmt_display:
                    is_display_block = True
                    edit_attrib_deciseconds = 200
                    continue
                else:
                    edit_attrib_deciseconds += 1
                    time.sleep(.1)
            except:
                logging.error("edit attribute dialog display never shifted to block\n")
                return False

        # Resetting edit attribute seconds variable for use later in this method
        edit_attrib_deciseconds = 0

        # Wait to make sure the driver finds the edit attribute dialog
        try:
            edit_attribute_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Edit Attribute dialog not found")
            logging.error("Edit Attribute dialog not found", exc_info=True)
            return False

        try:
            # Find the edit attribute dialog
            edit_attribute_dialog = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the edit attribute dialog")
            logging.error("WebDriver was not able to locate the edit attribute dialog", exc_info=True)
            return False

        # Setting loop variable to handle wait
        edit_title_deciseconds = 0  # Loop variable

        # Waiting until the edit attribute title includes the text: Company Net Content
        while edit_title_deciseconds < 200:  # 20 seconds
            try:
                is_Company_net_content = "Company Net Content" in edit_attribute_dialog.text

                if is_Company_net_content:
                    edit_title_deciseconds = 200
                    continue
            except:
                edit_title_deciseconds += 1
                time.sleep(0.1)

        # Read the title of the edit attribute dialog to make sure we are about to edit the correct attribute
        edit_attribute_title = edit_attribute_dialog.text

        if edit_attribute_title != "Company Net Content":
            print("Error:: The attribute selected was not Company Net Content")
            logging.error("The attribute selected was not Company Net Content")
            return False

        # Wait to make sure the driver finds the cell edit iframe
        try:
            cell_edit_iframe_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe"))
        except Exception as e:
            print("Error:: Cell Edit iframe not found")
            logging.error("Cell Edit iframe not found", exc_info=True)
            return False

        # Switch to the cell edit iframe
        iframe = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe")
        web_driver.switch_to.frame(iframe)

        # Wait to make sure the driver finds the attribute value field
        try:
            attribute_value_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Attribute Value field not found")
            logging.error("Attribute Value field not found", exc_info=True)
            return False

        try:
            # Find the attribute value field
            attribute_value_field = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the attribute value field")
            logging.error("WebDriver was not able to locate the attribute value field", exc_info=True)
            return False

        try:
            # Clear the existing attribute value input text
            attribute_value_field.clear()

            # Enter the fixed attribute value
            attribute_value_field.send_keys(attrib_value)

        except:
            # Wait to make sure the driver finds the cancel button
            try:
                cancel_button_wait = WebDriverWait(web_driver, timeout=20).until(
                    lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
            except Exception as e:
                print("Error:: Cancel button not found")
                logging.error("Cancel button not found", exc_info=True)
                return False

            try:
                # Find the cancel button
                cancel_button = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
            except Exception as e:
                print("Error:: WebDriver was not able to locate the cancel button")
                logging.error("WebDriver was not able to locate the cancel button", exc_info=True)
                return False

            # Click the cancel button
            cancel_button.click()
            return False

        # Wait to make sure the driver finds the save button field
        try:
            save_button_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Save button not found")
            logging.error("Save button not found", exc_info=True)
            return False

        try:
            # Find the save button
            save_button = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the Save button")
            logging.error("WebDriver was not able to locate the Save button", exc_info=True)
            return False

        # Click the save button
        save_button.click()

        # Wait to make sure the driver finds the maingrid iframe
        try:
            maingrid_iframe_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe"))
        except Exception as e:
            print("Error:: maingrid iframe not found")
            logging.error("maingrid iframe not found", exc_info=True)
            return False

        # Switch to the main grid iframe
        iframe = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe")
        web_driver.switch_to.frame(iframe)

        # Used in loop below
        is_edit_attrib_closed = False

        # Waiting to make sure the edit attribute dialog is removed from the DOM before proceeding
        while edit_attrib_deciseconds < 200:  # 20 seconds
            # For this situation the except code block is the outcome we're looking for
            try:
                # Check to see if web driver finds the element
                edit_attrib_element = web_driver.find_element(By.XPATH, edit_attrib_path)
                edit_attrib_deciseconds += 1
                time.sleep(.1)
            except:
                is_edit_attrib_closed = True
                edit_attrib_deciseconds = 200
                continue

        # When function completes without any errors
        if is_edit_attrib_closed:
            return True
        else:
            print("Error:: Encountered problem with the update_Company_net_content function.\n")
            logging.error("The edit attribute dialog never left the DOM")
            return False

    except Exception as e:
        print("Error:: Exception encountered when attempting to update the Company net content")
        logging.error("Exception encountered when attemtping to update the Company net content", exc_info=True)
        return False


def update_blank_net_content(attrib_value, web_driver):
    """
    This function is designed to update blank values in the Net Content field based on the Fixed value from the Company Net
    Content field.  This provides an opportunity to clear even more validation errors in Net Content for products where
    the Net Content was blank but the Company Net Content was not.
    ----------------------------------------------------------
    Based on the following requirements from the PIM Business Analyst:
    * Maximum Number of Characters = 9
    * No more than 2 digits to the right of the decimal
    * Should not be blank
    ---------------------
    * All values with more than 2 decimal digits should be modified to include only 2
    * Blank values should be modified to -1 to show that the program has already checked the value
    """

    try:

        edit_attrib_path = "//div[@aria-labelledby='PLACEHOLDER']"

        # Wait to make sure the driver finds the edit attribute dialog
        try:
            edit_attribute_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, edit_attrib_path))
        except Exception as e:
            print("Error:: Edit Attribute dialog not found")
            logging.error("Edit Attribute dialog not found", exc_info=True)
            return False

        # Setting loop variable to handle wait
        edit_attrib_deciseconds = 0  # Loop variable

        # Waiting until the edit attribute dialog display switches to block in the DOM
        while edit_attrib_deciseconds < 200:  # 20 seconds
            try:

                # Locate the edit attribute dialog element
                edit_attrib_element = web_driver.find_element(By.XPATH, edit_attrib_path)
                edit_elmt_display = edit_attrib_element.get_dom_attribute("style")

                if "display: block" in edit_elmt_display:
                    is_display_block = True
                    edit_attrib_deciseconds = 200
                    continue
                else:
                    edit_attrib_deciseconds += 1
                    time.sleep(.1)
            except:
                logging.error("edit attribute dialog display never shifted to block\n")
                return False

        # Resetting edit attribute seconds variable for use later in this method
        edit_attrib_deciseconds = 0

        # Wait to make sure the driver finds the edit attribute dialog
        try:
            edit_attribute_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Edit Attribute dialog not found")
            logging.error("Edit Attribute dialog not found", exc_info=True)
            return False

        try:
            # Find the edit attribute dialog
            edit_attribute_dialog = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the edit attribute dialog")
            logging.error("WebDriver was not able to locate the edit attribute dialog", exc_info=True)
            return False

        # Setting loop variable to handle wait
        edit_title_deciseconds = 0  # Loop variable

        # Waiting until the edit attribute title includes the text: Net Content
        while edit_title_deciseconds < 200:  # 20 seconds
            try:
                is_net_content = "Net Content" in edit_attribute_dialog.text

                if is_net_content:
                    edit_title_deciseconds = 200
                    continue
            except:
                edit_title_deciseconds += 1
                time.sleep(0.1)

        # Read the title of the edit attribute dialog to make sure we are about to edit the correct attribute
        edit_attribute_title = edit_attribute_dialog.text

        if edit_attribute_title != "Net Content":
            print("Error:: The attribute selected was not Net Content")
            logging.error("The attribute selected was not Net Content")
            return False

        # Wait to make sure the driver finds the cell edit iframe
        try:
            cell_edit_iframe_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe"))
        except Exception as e:
            print("Error:: Cell Edit iframe not found")
            logging.error("Cell Edit iframe not found", exc_info=True)
            return False

        # Switch to the cell edit iframe
        iframe = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe")
        web_driver.switch_to.frame(iframe)

        # Wait to make sure the driver finds the attribute value field
        try:
            attribute_value_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Attribute Value field not found")
            logging.error("Attribute Value field not found", exc_info=True)
            return False

        try:
            # Find the attribute value field
            attribute_value_field = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the attribute value field")
            logging.error("WebDriver was not able to locate the attribute value field", exc_info=True)
            return False

        try:
            # Clear the existing attribute value input text
            attribute_value_field.clear()

            # Enter the fixed attribute value
            attribute_value_field.send_keys(attrib_value)

        except:
            # Wait to make sure the driver finds the cancel button
            try:
                cancel_button_wait = WebDriverWait(web_driver, timeout=20).until(
                    lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
            except Exception as e:
                print("Error:: Cancel button not found")
                logging.error("Cancel button not found", exc_info=True)
                return False

            try:
                # Find the cancel button
                cancel_button = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
            except Exception as e:
                print("Error:: WebDriver was not able to locate the cancel button")
                logging.error("WebDriver was not able to locate the cancel button", exc_info=True)
                return False

            # Click the cancel button
            cancel_button.click()
            return False

        # Wait to make sure the driver finds the save button field
        try:
            save_button_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']"))
        except Exception as e:
            print("Error:: Save button not found")
            logging.error("Save button not found", exc_info=True)
            return False

        try:
            # Find the save button
            save_button = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']")
        except Exception as e:
            print("Error:: WebDriver was not able to locate the Save button")
            logging.error("WebDriver was not able to locate the Save button", exc_info=True)
            return False

        # Click the save button
        save_button.click()

        # Wait to make sure the driver finds the maingrid iframe
        try:
            maingrid_iframe_wait = WebDriverWait(web_driver, timeout=20).until(
                lambda document: document.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe"))
        except Exception as e:
            print("Error:: maingrid iframe not found")
            logging.error("maingrid iframe not found", exc_info=True)
            return False

        # Switch to the main grid iframe
        iframe = web_driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe")
        web_driver.switch_to.frame(iframe)

        # Used in loop below
        is_edit_attrib_closed = False

        # Waiting to make sure the edit attribute dialog is removed from the DOM before proceeding
        while edit_attrib_deciseconds < 200:  # 20 seconds
            # For this situation the except code block is the outcome we're looking for
            try:
                # Check to see if web driver finds the element
                edit_attrib_element = web_driver.find_element(By.XPATH, edit_attrib_path)
                edit_attrib_deciseconds += 1
                time.sleep(.1)
            except:
                is_edit_attrib_closed = True
                edit_attrib_deciseconds = 200
                continue

        # When function completes without any errors
        if is_edit_attrib_closed:
            return True
        else:
            print("Error:: Encountered problem with the update_net_content function.\n")
            logging.error("The edit attribute dialog never left the DOM")
            return False

    except Exception as e:
        print("Error:: Exception encountered when attempting to update the net content")
        logging.error("Exception encountered when attemtping to update the net content", exc_info=True)
        return False


def navigate_to_nextpage(web_driver, current_pg):

    # Using page_number_input instead of the next element because the next element is not directly interactable
    try:
        # Wait to make sure the driver finds the page_number_input field
        page_number_input_wait = WebDriverWait(web_driver, timeout=20).until(
            lambda document: document.find_element(By.XPATH, "//input[@class='PLACEHOLDER']"))
    except Exception as e:
        print("Error:: Page Number Input field not found")
        logging.error("Page Number Input field not found", exc_info=True)
        return False

    # Brief wait to account for browser latency
    time.sleep(1)

    current_pg_int = int(current_pg)
    next_pg = current_pg_int + 1

    try:
        # Retrieve updated data for the page number input field
        page_number_input = web_driver.find_element(By.XPATH, "//input[@class='PLACEHOLDER']")

        # Clear the page_number_input text
        page_number_input.clear()
        # Brief delay to account for browser latency
        time.sleep(.5)
        # Enter the number for the next page
        page_number_input.send_keys(str(next_pg))
        # Brief delay to account for browser latency
        time.sleep(.5)
        # Hit the RETURN key to navigate to the next page of search results
        page_number_input.send_keys(Keys.RETURN)

    except Exception as e:
        print("Error:: Encountered problem with the navigate_to_nextpage function.\n")
        logging.error("Exception occurred", exc_info=True)
        return False

    # When function completes without any errors
    return True


def print_activity_summary():

    # Call the flash_window function
    flash_window()

    logging.info("\n\n                               *********************************\n"
                 "                               *        Activity Summary        *\n"
                 "                               *********************************")

    print("\n\n*********************************\n"
          "*        Activity Summary       *\n"
          "*********************************\n"
          "                               ")
    logging.info(
        "\n\n                               Total Products Reviewed: " + str(items_reviewed_counter))
    print("Total Products Reviewed:  " + str(items_reviewed_counter))

    logging.info(
        "\n\n                               Total Products Fixed:    " + str(items_fixed_counter))
    print("Total Products Corrected: " + str(items_fixed_counter))

    logging.info(
        "\n\n                               Fixed " + str(errors_fixed_counter))
    print("Total Errors Corrected:   " + str(errors_fixed_counter) + "\n")
    print("*********************************\n")


def save_and_quit():
    """
    Saves program activity to files and exits the program. Saves data for the records which were reviewed to one file,
    and saves data for the records which were corrected to another file.
    There is also a hotkey set up to call this method even while the program is still running to make sure the
    activity gets saved to file if the program needs to be interrupted.
    *******************
    *  Hotkey: <Alt+C>  *
    *******************
    """

    try:
        print("\n\n*********************************\n"
              "Preparing data to be saved to file...\n")
    except Exception as e:
        logging.error("Exception occurred while attempting to save and exit.",
                      exc_info=True)
        print("Error:: Exception occurred while attempting to save and exit.")

    try:
        try:
            # Calculating number of products reviewed and number of products fixed using global list variables
            reviewed_count = int(len(reviewed) / 6)    # 7 pieces of data were added per product
            fixed_count = int(len(fixed) / 11)          # 13 pieces of data were added per product

            # Create 2D numpy arrays
            reviewed_items = numpy.array(reviewed, dtype=str).reshape(reviewed_count, 6)
            fixed_items = numpy.array(fixed, dtype=str).reshape(fixed_count, 11)

            print("Data preparation completed.  Saving data to files...\n")

        except Exception as e:
            logging.error("Exception occurred while attempting to create numpy arrays for reviewed and corrected items.",
                          exc_info=True)
            print("Error:: Exception occurred while attempting to create numpy arrays for reviewed and corrected items.")

        # Capture current datatime
        file_dtnow = datetime.datetime.now()

        # Format datetime
        file_datetime = file_dtnow.strftime("%m.%d.%Y_%H.%M.%S")

        # Define filenames for each file
        summary_filename = file_datetime + '_data_cleanup_summary.txt'
        reviewed_filename = file_datetime + '_data_cleanup_reviewed.csv'
        corrected_filename = file_datetime + '_data_cleanup_fixed.csv'

        # Save the activity summary data to a txt file
        try:
            with open(summary_path + "/" + summary_filename, 'w') as summary_file:
                summary_file.write(
                    "PIM Data Cleanup Activity Summary\n"
                    "----------------------------------\n"
                    "Total Products Reviewed: " + str(items_reviewed_counter) + "\n"
                    "Total Products Fixed:    " + str(items_fixed_counter) + "\n"
                    "Total Errors Fixed:      " + str(errors_fixed_counter) + "\n"
                )
        except Exception as e:
            logging.error("Exception occurred while attempting to save activity summary data to txt file.",
                          exc_info=True)
            print(
                "Error:: Exception occurred while attempting to save activity summary data to txt file.")

        # Save the reviewed items data to a csv file
        try:
            numpy.savetxt(reviewed_path + "/" + reviewed_filename, reviewed_items, fmt='%s', delimiter=",")
            print("Reviewed item data saved successfully.  Now saving corrected item data...\n")
        except Exception as e:
            logging.error("Exception occurred while attempting to save reviewed item data to csv file.", exc_info=True)
            print(
                "Error:: Exception occurred while attempting to save reviewed item data to csv file.")

        # Save the fixed items data to a csv file
        try:
            numpy.savetxt(corrected_path + "/" + corrected_filename, fixed_items, fmt='%s', delimiter=",")
            print("Corrected item data saved successfully.\n")
            print("Please close the application.")

            # Effectively freeze the program to give user an opportunity to review console info before closing
            freeze_event.set()


        except Exception as e:
            logging.error("Exception occurred while attempting to save reviewed item data to csv file.", exc_info=True)
            print(
                "Error:: Exception occurred while attempting to save reviewed item data to csv file.")
            print("Please close the application.")

            # Effectively freeze the program to give user an opportunity to review console info before closing
            freeze_event.set()

        # Effectively freeze the program to give user an opportunity to review console info before closing
        freeze_event.set()

    except Exception as e:
        logging.error("Exception occurred while attempting to create reviewed and corrected files.", exc_info=True)
        print("Error:: Exception occurred while attempting to create reviewed and corrected files.")
        print("Please close the application.")

        # Effectively freeze the program to give user an opportunity to review console info before closing
        freeze_event.set()

    # Effectively freeze the program to give user an opportunity to review console info before closing
    freeze_event.set()


def main():
    """
    The Main Method
    """

    # Setting hotkey to call the save_and_exit method
    keyboard.add_hotkey('alt+c', save_and_quit)

    try:
        # Initializing program
        init_logger()
        get_user_input_prerequisites()
        driver = init_webdriver()

        # Initializing window flash to let user know when main program loop is completed
        kernel32.GetConsoleWindow.restype = wintypes.HWND
        user32.FlashWindowEx.argtypes = (ctypes.POINTER(WindowFlash),)

        # Switch to the main grid iframe
        iframe = driver.find_element(By.XPATH, "//*[@id='PLACEHOLDER']/iframe")
        driver.switch_to.frame(iframe)

        paging_info = parse_paging_info(driver)
        current_row_on_page = 1  # Variable to track which row on the page we're reviewing
        total_records = get_total_records(paging_info)
        total_pages = get_total_pages(total_records)

        # Print and log initial values used for main program loop
        print_initial_values(total_pages, total_records)

        # Initializing program flow variables for main program loop
        is_finished = False
        product_updated = False
        first_iteration = True
        number_of_hiccups = 0

        # Bringing global counters into the main method
        global items_reviewed_counter
        global items_fixed_counter
        global errors_fixed_counter

        # Initializing attribute value variables for main program loop
        original_manufacturer_number = ""
        fixed_manufacturer_number = ""
        original_start_availability = ""
        fixed_start_availability = ""
        original_master_gtin = ""
        fixed_master_gtin = ""
        original_net_content = ""
        fixed_net_content = ""
        original_Company_net_content = ""
        fixed_Company_net_content = ""

        '''
        The Main Program Loop
        ---------------------
        
        General flow: 
        
            - Get original values for the product
            - Check values to see if any are invalid
            - For any invalid values:
                - calculate valid value 
                - fix it
                
        '''

        while number_of_hiccups < 10 and not is_finished and not freeze_event.is_set():

            # Not triggered first time through loop, but IS triggered every other time
            if not first_iteration:

                # Flip first_iteration to False so the program knows that any iteration after this is NOT the first
                first_iteration = False

                # Make sure the main grid has finished loading
                if not check_lui_maingrid(driver):
                    # Increment the number of hiccups and try to overcome the system hiccup without crashing
                    number_of_hiccups += 1
                    print("\nEncountered hiccup in the system.\n"
                          "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                    flash_window()
                    continue

            '''
            Collecting data from the current product record
            '''

            # Get the element id for the current table row
            current_row_id = get_row_id(driver, current_row_on_page)

            if current_row_id == "error":  # If there was an error
                # Increment the number of hiccups and try to overcome the system hiccup without crashing
                number_of_hiccups += 1
                print("\nEncountered hiccup in the system.\n"
                      "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                flash_window()
                continue

            # Get the Company product number from the current record
            Company_product_number = get_Company_prod_number(driver, current_row_id)
            print("Company Product Number: " + str(Company_product_number))

            if Company_product_number == "error":
                # Increment the number of hiccups and try to overcome the system hiccup without crashing
                number_of_hiccups += 1
                print("\nEncountered hiccup in the system.\n"
                      "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                flash_window()
                continue

            # Get the manufacturer number value from the current record
            original_manufacturer_number = get_manufacturer_number(driver, current_row_id)

            if original_manufacturer_number == "error":
                # Increment the number of hiccups and try to overcome the system hiccup without crashing
                number_of_hiccups += 1
                print("\nEncountered hiccup in the system.\n"
                      "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                flash_window()
                continue

            # Get the start availability date time value from the current record
            original_start_availability = get_start_availability(driver, current_row_id)

            if original_start_availability == "error":
                # Increment the number of hiccups and try to overcome the system hiccup without crashing
                number_of_hiccups += 1
                print("\nEncountered hiccup in the system.\n"
                      "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                flash_window()
                continue

            # Get the master gtin value from the current record
            original_master_gtin = get_master_gtin(driver, current_row_id)

            if original_master_gtin == "error":
                # Increment the number of hiccups and try to overcome the system hiccup without crashing
                number_of_hiccups += 1
                print("\nEncountered hiccup in the system.\n"
                      "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                flash_window()
                continue

            # Get the net content value from the current record
            original_net_content = get_net_content(driver, current_row_id)

            if original_net_content == "error":
                # Increment the number of hiccups and try to overcome the system hiccup without crashing
                number_of_hiccups += 1
                print("\nEncountered hiccup in the system.\n"
                      "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                flash_window()
                continue

            # Get the Company net content value from the current record
            original_Company_net_content = get_Company_net_content(driver, current_row_id)

            if original_Company_net_content == "error":
                # Increment the number of hiccups and try to overcome the system hiccup without crashing
                number_of_hiccups += 1
                print("\nEncountered hiccup in the system.\n"
                      "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                flash_window()
                continue

            '''
            Checking data for validity
            '''

            # Check to see if value in manufacturer number field is valid
            manufacturer_number_valid = is_manufacturer_number_valid(original_manufacturer_number)

            # If error
            if manufacturer_number_valid == "error":
                # Increment the number of hiccups and try to overcome the system hiccup without crashing
                number_of_hiccups += 1
                print("\nEncountered hiccup in the system.\n"
                      "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                flash_window()
                continue

            # Check to see if value in start availability date time field is valid
            start_availability_valid = is_start_availability_valid(original_start_availability)

            # If error
            if start_availability_valid == "error":
                # Increment the number of hiccups and try to overcome the system hiccup without crashing
                number_of_hiccups += 1
                print("\nEncountered hiccup in the system.\n"
                      "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                flash_window()
                continue

            # Doublecheck start availability date time by checking the edit attribute dialog
            if not start_availability_valid:

                # capture the value from the edit attribute dialog
                original_start_availability = get_start_availability_dialog(driver, current_row_id)

                # If error
                if start_availability_valid == "error":
                    # Increment the number of hiccups and try to overcome the system hiccup without crashing
                    number_of_hiccups += 1
                    print("\nEncountered hiccup in the system.\n"
                          "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                    flash_window()
                    continue

                # Check to see if the value from the edit attribute dialog is valid
                start_availability_valid = doublecheck_start_availability(original_start_availability)

                # If error
                if start_availability_valid == "error":
                    # Increment the number of hiccups and try to overcome the system hiccup without crashing
                    number_of_hiccups += 1
                    print("\nEncountered hiccup in the system.\n"
                          "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                    flash_window()
                    continue

            # Check to see if value in master gtin field is valid
            master_gtin_valid = is_master_gtin_valid(original_master_gtin)

            # If error
            if master_gtin_valid == "error":
                # Increment the number of hiccups and try to overcome the system hiccup without crashing
                number_of_hiccups += 1
                print("\nEncountered hiccup in the system.\n"
                      "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                flash_window()
                continue

            # Check to see if value in net content field is blank
            net_content_blank = is_net_content_blank(original_net_content)

            # If error
            if net_content_blank == "error":
                # Increment the number of hiccups and try to overcome the system hiccup without crashing
                number_of_hiccups += 1
                print("\nEncountered hiccup in the system.\n"
                      "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                flash_window()
                continue

            # Check to see if value in Company net content field is valid
            Company_net_content_valid = is_Company_net_content_valid(original_Company_net_content)

            # If error
            if Company_net_content_valid == "error":
                # Increment the number of hiccups and try to overcome the system hiccup without crashing
                number_of_hiccups += 1
                print("\nEncountered hiccup in the system.\n"
                      "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                flash_window()
                continue

            '''
            Calculating valid values
            '''

            # Calculate the valid manufacturer number if value is invalid
            if not manufacturer_number_valid:

                # calculate the correct value
                fixed_manufacturer_number = calculate_manufacturer_number(original_manufacturer_number)

                # If error
                if fixed_manufacturer_number == "error":
                    # Increment the number of hiccups and try to overcome the system hiccup without crashing
                    number_of_hiccups += 1
                    print("\nEncountered hiccup in the system.\n"
                          "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                    flash_window()
                    continue

            # Calculate the valid master gtin if value is invalid
            if not master_gtin_valid:

                # calculate the correct value
                fixed_master_gtin = calculate_master_gtin(original_master_gtin)

                # If error
                if fixed_master_gtin == "error":
                    # Increment the number of hiccups and try to overcome the system hiccup without crashing
                    number_of_hiccups += 1
                    print("\nEncountered hiccup in the system.\n"
                          "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                    flash_window()
                    continue

            # Calculate the valid Company net content if value is invalid
            if not Company_net_content_valid:

                # calculate the correct value
                fixed_Company_net_content = calculate_Company_net_content(driver, current_row_id)

                # If error
                if fixed_Company_net_content == "error":
                    # Increment the number of hiccups and try to overcome the system hiccup without crashing
                    number_of_hiccups += 1
                    print("\nEncountered hiccup in the system.\n"
                          "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                    flash_window()
                    continue

            '''
            Updating products to fix invalid data
            '''

            # Manage program flow related to threading in case the alt+c hotkey is pressed
            if not freeze_event.is_set():

                # Set product_updated to False, will flip to True if product is updated in sections below
                product_updated = False

                if not manufacturer_number_valid:

                    # Double-click the manufacturer number field to open the edit attribute dialog
                    if not click_manufacturer_number(driver, current_row_id):
                        # Increment the number of hiccups and try to overcome the system hiccup without crashing
                        number_of_hiccups += 1
                        print("\nEncountered hiccup in the system.\n"
                              "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                        flash_window()
                        continue

                    # Update the manufacturer number for the selected record
                    if not update_manufacturer_number(fixed_manufacturer_number, driver):
                        # Increment the number of hiccups and try to overcome the system hiccup without crashing
                        number_of_hiccups += 1
                        print("\nEncountered hiccup in the system.\n"
                              "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                        flash_window()
                        continue

                    # Make sure the main grid has finished loading
                    if not check_lui_maingrid_click(driver):
                        # Increment the number of hiccups and try to overcome the system hiccup without crashing
                        number_of_hiccups += 1
                        print("\nEncountered hiccup in the system.\n"
                              "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                        flash_window()
                        continue

                    product_updated = True

                    # Update error counter
                    errors_fixed_counter += 1

                    print("Original Manufacturer Number: " + original_manufacturer_number)
                    print("Corrected Manufacturer Number: " + fixed_manufacturer_number)

            # Manage program flow related to threading in case the alt+c hotkey is pressed
            if not freeze_event.is_set():

                if not start_availability_valid:

                    fixed_start_availability = ""

                    # Double-click the start availability date time field to open the edit attribute dialog
                    if not click_start_availability(driver, current_row_id):
                        # Increment the number of hiccups and try to overcome the system hiccup without crashing
                        number_of_hiccups += 1
                        print("\nEncountered hiccup in the system.\n"
                              "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                        flash_window()
                        continue

                    # Update the start availability date time for the selected record
                    if not update_start_availability(driver):
                        # Increment the number of hiccups and try to overcome the system hiccup without crashing
                        number_of_hiccups += 1
                        print("\nEncountered hiccup in the system.\n"
                              "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                        flash_window()
                        continue

                    # Make sure the main grid has finished loading
                    if not check_lui_maingrid_click(driver):
                        # Increment the number of hiccups and try to overcome the system hiccup without crashing
                        number_of_hiccups += 1
                        print("\nEncountered hiccup in the system.\n"
                              "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                        flash_window()
                        continue

                    product_updated = True

                    # Update error counter
                    errors_fixed_counter += 1

                    print("Original Start Availability Date Time: " + original_start_availability)
                    print("Corrected Start Availability Date Time: " + fixed_start_availability)

            # Manage program flow related to threading in case the alt+c hotkey is pressed
            if not freeze_event.is_set():

                if not master_gtin_valid:

                    # Double-click the master gtin field to open the edit attribute dialog
                    if not click_master_gtin(driver, current_row_id):
                        # Increment the number of hiccups and try to overcome the system hiccup without crashing
                        number_of_hiccups += 1
                        print("\nEncountered hiccup in the system.\n"
                              "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                        flash_window()
                        continue

                    # Update the master gtin for the selected record
                    if not update_master_gtin(fixed_master_gtin, driver):
                        # Increment the number of hiccups and try to overcome the system hiccup without crashing
                        number_of_hiccups += 1
                        print("\nEncountered hiccup in the system.\n"
                              "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                        flash_window()
                        continue

                    # Make sure the main grid has finished loading
                    if not check_lui_maingrid_click(driver):
                        # Increment the number of hiccups and try to overcome the system hiccup without crashing
                        number_of_hiccups += 1
                        print("\nEncountered hiccup in the system.\n"
                              "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                        flash_window()
                        continue

                    product_updated = True

                    # Update error counter
                    errors_fixed_counter += 1

                    print("Original Master GTIN: " + original_master_gtin)
                    print("Corrected Master GTIN: " + fixed_master_gtin)

            # Manage program flow related to threading in case the alt+c hotkey is pressed
            if not freeze_event.is_set():

                if not Company_net_content_valid:

                    # Double-click the Company net content field to open the edit attribute dialog
                    if not click_Company_net_content(driver, current_row_id):
                        # Increment the number of hiccups and try to overcome the system hiccup without crashing
                        number_of_hiccups += 1
                        print("\nEncountered hiccup in the system.\n"
                              "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                        flash_window()
                        continue

                    # Update the Company net content for the selected record
                    if not update_Company_net_content(fixed_Company_net_content, driver):
                        # Increment the number of hiccups and try to overcome the system hiccup without crashing
                        number_of_hiccups += 1
                        print("\nEncountered hiccup in the system.\n"
                              "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                        flash_window()
                        continue

                    # Make sure the main grid has finished loading
                    if not check_lui_maingrid_click(driver):
                        # Increment the number of hiccups and try to overcome the system hiccup without crashing
                        number_of_hiccups += 1
                        print("\nEncountered hiccup in the system.\n"
                              "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                        flash_window()
                        continue

                    product_updated = True

                    # Update error counter
                    errors_fixed_counter += 1

                    print("Original Company Net Content: " + original_Company_net_content)
                    print("Corrected Company Net Content: " + fixed_Company_net_content)

            # Manage program flow related to threading in case the alt+c hotkey is pressed
            if not freeze_event.is_set():

                if net_content_blank:

                    # Check if the Company net content is valid so we can use that
                    if is_Company_net_content_valid(original_Company_net_content) or is_Company_net_content_valid(fixed_Company_net_content):

                        if is_Company_net_content_valid(original_Company_net_content):
                            fixed_net_content = original_Company_net_content
                        elif is_Company_net_content_valid(fixed_Company_net_content):
                            fixed_net_content = fixed_Company_net_content

                        # Double-click the net content field to open the edit attribute dialog
                        if not click_net_content(driver, current_row_id):
                            # Increment the number of hiccups and try to overcome the system hiccup without crashing
                            number_of_hiccups += 1
                            print("\nEncountered hiccup in the system.\n"
                                  "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                            flash_window()
                            continue

                        # Update the net content for the selected record
                        if not update_blank_net_content(fixed_net_content, driver):
                            # Increment the number of hiccups and try to overcome the system hiccup without crashing
                            number_of_hiccups += 1
                            print("\nEncountered hiccup in the system.\n"
                                  "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                            flash_window()
                            continue

                        # Make sure the main grid has finished loading
                        if not check_lui_maingrid_click(driver):
                            # Increment the number of hiccups and try to overcome the system hiccup without crashing
                            number_of_hiccups += 1
                            print("\nEncountered hiccup in the system.\n"
                                  "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                            flash_window()
                            continue

                        product_updated = True

                        # Update error counter
                        errors_fixed_counter += 1

                        print("Original Net Content: " + original_net_content)
                        print("Corrected Net Content: " + fixed_net_content)

                    # Check if Company net content was fixed to -1 so we can use that
                    elif fixed_Company_net_content == '-1':

                        fixed_net_content = fixed_Company_net_content

                        # Double-click the net content field to open the edit attribute dialog
                        if not click_net_content(driver, current_row_id):
                            # Increment the number of hiccups and try to overcome the system hiccup without crashing
                            number_of_hiccups += 1
                            print("\nEncountered hiccup in the system.\n"
                                  "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                            flash_window()
                            continue

                        # Update the net content for the selected record
                        if not update_blank_net_content(fixed_net_content, driver):
                            # Increment the number of hiccups and try to overcome the system hiccup without crashing
                            number_of_hiccups += 1
                            print("\nEncountered hiccup in the system.\n"
                                  "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                            flash_window()
                            continue

                        # Make sure the main grid has finished loading
                        if not check_lui_maingrid_click(driver):
                            # Increment the number of hiccups and try to overcome the system hiccup without crashing
                            number_of_hiccups += 1
                            print("\nEncountered hiccup in the system.\n"
                                  "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                            flash_window()
                            continue

                        product_updated = True

                        # Update error counter
                        errors_fixed_counter += 1

                        print("Original Net Content: " + original_net_content)
                        print("Corrected Net Content: " + fixed_net_content)

                    # Check if Company net content was fixed to a valid value (not -1) so we can use that
                    elif original_net_content != '-1':

                        fixed_net_content = '-1'

                        # Double-click the net content field to open the edit attribute dialog
                        if not click_net_content(driver, current_row_id):
                            # Increment the number of hiccups and try to overcome the system hiccup without crashing
                            number_of_hiccups += 1
                            print("\nEncountered hiccup in the system.\n"
                                  "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                            flash_window()
                            continue

                        # Update the net content for the selected record
                        if not update_blank_net_content(fixed_net_content, driver):
                            # Increment the number of hiccups and try to overcome the system hiccup without crashing
                            number_of_hiccups += 1
                            print("\nEncountered hiccup in the system.\n"
                                  "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                            flash_window()
                            continue

                        # Make sure the main grid has finished loading
                        if not check_lui_maingrid_click(driver):
                            # Increment the number of hiccups and try to overcome the system hiccup without crashing
                            number_of_hiccups += 1
                            print("\nEncountered hiccup in the system.\n"
                                  "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                            flash_window()
                            continue

                        product_updated = True

                        # Update error counter
                        errors_fixed_counter += 1

                        print("Original Net Content: " + original_net_content)
                        print("Corrected Net Content: " + fixed_net_content)

                # Add the Company product number and net content for the current item to the main reviewed array
                current_reviewed_data = [Company_product_number, original_manufacturer_number, original_start_availability,
                                         original_master_gtin, original_net_content,
                                         original_Company_net_content]
                reviewed.extend(current_reviewed_data)

            if product_updated:

                # Manage program flow related to threading in case the alt+c hotkey is pressed
                if not freeze_event.is_set():

                    # Update items fixed counter
                    items_fixed_counter += 1
                    print("\nTotal Products Updated: " + str(items_fixed_counter))

                    # Print errors corrected counter
                    print("Total Errors Corrected: " + str(errors_fixed_counter) + "\n")

                # Add the Company product number, original attribute values, and fixed values to the main fixed array
                current_fixed_data = [Company_product_number, original_manufacturer_number, fixed_manufacturer_number,
                                      original_start_availability, fixed_start_availability, original_master_gtin,
                                      fixed_master_gtin, original_net_content, fixed_net_content,
                                      original_Company_net_content, fixed_Company_net_content]
                fixed.extend(current_fixed_data)

                # Reset all the fixed value variables
                fixed_manufacturer_number = ""
                fixed_start_availability = ""
                fixed_master_gtin = ""
                fixed_net_content = ""
                fixed_Company_net_content = ""

            # Manage program flow related to threading in case the alt+c hotkey is pressed
            if not freeze_event.is_set():

                # Update items reviewed counter
                items_reviewed_counter += 1

            # Get the updated paging info
            paging_info = parse_paging_info(driver)
            first_record_on_page = get_first_record_on_page(paging_info)
            last_record_on_page = get_last_record_on_page(paging_info)
            current_record = first_record_on_page + current_row_on_page  # Variable to track which record we're on
            current_page = get_current_page(first_record_on_page)
            total_records = get_total_records(paging_info)

            # Check to see if the program has reviewed all records
            if current_record == total_records:

                is_finished = True
                print("\n===============================\n"
                      "End of product records.\n"
                      "===============================\n")
                logging.info("\n\n==========================\n"
                             "End of product records.\n"
                             "==========================\n")

                # Exit the while loop
                break

            # Else if there are still records to review
            else:

                # Check to see if on the last record of the page
                if current_record > last_record_on_page and not freeze_event.is_set():

                    print("\nNavigating to next page...\n")
                    print("Hit Alt+C at any time to save program activity to file and exit.\n")

                    # Navigate to the next page of search results
                    if not navigate_to_nextpage(driver, current_page):
                        # Increment the number of hiccups and try to overcome the system hiccup without crashing
                        number_of_hiccups += 1
                        print("\nEncountered hiccup in the system.\n"
                              "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                        flash_window()
                        continue

                    # Make sure the main grid has finished loading
                    if not check_lui_maingrid(driver):
                        # Increment the number of hiccups and try to overcome the system hiccup without crashing
                        number_of_hiccups += 1
                        print("\nEncountered hiccup in the system.\n"
                              "Number of system hiccups encountered so far: " + str(number_of_hiccups) + "\n")
                        flash_window()
                        continue

                    # Reset the current_row_on_page variable
                    current_row_on_page = 1

                    # Increment the current_page variable
                    current_page += 1

                # If not the last record on the page, advance to the next record on the page
                elif current_record <= last_record_on_page:

                    # Increment current record on page
                    current_row_on_page += 1

        # Flash the tray icon in the taskbar to let user know the program is finished
        flash_window()

        # Only call the save_and_quit function if the program reached the end of the product records to review
        if not freeze_event.is_set():

            # Save program activity to files and exit the program
            save_and_quit()

        # Print the activity summary
        print_activity_summary()

        input("\nPress ENTER to Exit the PIM Data Cleanup program.\n\n")

    except Exception as e:
        flash_window()
        logging.error("Exception occurred during main loop.", exc_info=True)
        print("Exception occurred during main program loop.")
        time.sleep(3)
        save_and_quit()


if __name__ == "__main__":
    main()

# If program flow does not get caught by exceptions to trigger save_and_quit method, program exits here
sys.exit()
