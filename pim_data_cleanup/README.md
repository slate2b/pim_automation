# PIM Data Cleanup Automation Program
A program designed to automate data cleanup activities which must be performed through a web application's UI through a browser.  

This program was written in Python and utilizes Selenium WebDriver to interact with a web-based enterprise Product Information Management (PIM) system.

<h2>Background</h2>

The PIM system which this Data Cleanup program was designed for contains product records with many data attributes.  It uses validation rules for these attributes to help ensure that the product data in the system is valid.  However, due to problems with data migration and due to changes in the validation rules, many of the attribute values on these product records are now triggering validation errors.  Until the validation errors on a product record are cleared, that product record cannot be updated.  This causes significant issues when we need to update the product data on thousands of products at a time.

Unfortunately, due to system restrictions, performing data cleanup on product records through SQL scripts and other such methods was not an option.  In order to clear these errors, we had to go through a browser utilizing the PIM system's UI to correct the invalid values one product at a time.  With 1,000,000+ products in the system, this posed a significant problem.

This PIM Data Cleanup program was created to solve those problems.

<br>

https://github.com/slate2b/pim_automation_w_selenium/assets/88697660/b4621d83-b2ae-47cc-bd17-f3e26347001d

<h2>Functionality</h2>

The PIM Data Cleanup program automatically cycles through product records, identifies validation errors, dynamically calculates valid values, and then updates the product records, thereby clearing the validation errors and removing the major roadblocks to updating product records in the system.

Specifically, the program performs the following actions as it cycles through each product record:

* Reads the existing data in the following fields:
    * Manufacturer Number
    * Start Availability Datetime
    * Master GTIN
    * Net Content
    * USF Net Content
* Checks the existing data against validation logic
* Calculates the correct value(s)
* Updates the product record

<h2>Performance</h2>

The program is fully automated and can run in the background as a user completes other activities on their machine.  Performance is tied almost entirely to loading time for various dialogs and grid refreshes in the web-based PIM system itself.  Since errors must be corrected individually, each correction introduces wait time for the PIM system to communicate with its backend and database, verify the update, and then close the dialog / update the cell.  To mitigate the impact of these unavoidable delays from the PIM system, this Data Cleanup program was created using dynamic selenium waits and some custom waits to optimize performance in the Data Cleanup program itself.

Here are the average performance figures:

  * Reviews 877 records per hour
  * Clears 506 errors per hour

To put this in perspective, we often have to update product data on 10,000+ records at a time.  If a human user were manually reviewing 10,000 products with validation errors, it would take them approximately 20 hours to clear them.  However, if the PIM Data Cleanup program tackles those same 10,000 products, it would take less than 12 hours of automation processing with no need for any manual interaction from a human user.

  * Saves 20 Hours of Manual Work per 10,000 Products
  * Saves 200 Hours of Manual Work per 100,000 Products
  * Saves 2000 Hours of Manual Work per 1,000,000 Products

Note: The PIM System for which this program was created has a product catalog of 1,008,000+ products.

<h2>Accountability</h2>

The program was designed with multiple layers of safeguards in place to ensure that it would only interact with the correct web elements.  Basically, it looks before it leaps, not just once, but multiple times.  For example, if it expects to find a particular data field in one place and clicks on it to open a dialog, it checks the dialog title to make sure the title represents the actual attribute it's trying to interact with.  If the title doesn't match, it cancels the operation.

The program also creates and saves multiple files designed to provide a record of the program's activity each time it runs.

  * Saves a log file to capture a record of exceptions or other errors
  * Saves two csv files each time it runs, one containing data from the records reviewed and the other containing data from records corrected.  These files can be used to identify erroneous data loaded into the system in case the program fails to operate as expected, and also helps to identify the original data values so the product record can be returned to its original state.
  * Saves an activity summary to a text file each time it runs. The file contains stats for the number of products reviewed, number of products fixed, and the number of errors fixed.

<h2>Resilience</h2>

The program utilizes a combination of Selenium Waits and custom waits to dramatically reduce the likelihood of program crashes due to delays from page loads or brief network outages.  This increases the resilience of the program, making it highly resistant to exceptions and program crashes.

It also includes a hiccup-catching mechanism so that if the program exceeds the designated wait time for a particular interaction, the program simply increments its hiccup counter and restarts the current iteration of the main program loop.  This allows it to review the record it was in the middle of again without crashing.  I set it to save its activity files and stop processing if it encounters 10 hiccups because that would likely only happen if there were a significant issue like a prolonged loss of internet connectivity.

During testing and during implementation in production, this program was able to run without interruption for 48 hours straight on multiple occasions, only to stop when it reached the end of the record set it was reviewing.

<h2>Code Completeness and Implementation in Other Environments</h2>

This program was written to be used in an enterprise PIM system, so I have modified and/or removed some of the
particular DOM element references and other system-specific information for security purposes.

I am sharing this code in the hopes that the framework will be useful for someone who is looking for a similar
automation solution.

Anyone wishing to use a similar approach should keep in mind that the work of identifying the reference data
to help WebDriver find the elements you are trying to work with will still need to be performed, as well as
analyzing the state changes and data flow of the web application you are interacting with.
