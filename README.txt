FOIA REQUEST WEBSCRAPER

To start, extract the files of the zipped folder to a location of your choice. The required files are [FOIA]FOIA Requests.xlsx, gridView.xlsx, 
and [FOIA]WebCrawler.py. To use the FOIA Request Webscraper, ensure that Python 3.10 is downloaded (https://www.python.org/downloads/) and 
necessary modules are installed. To install modules execute the commands listed below in command prompt (Search for cmd and press enter).

py -m pip install pandas
py -m pip install requests
py -m pip install time
py -m pip install os

In case the error "'py' is not recognized as an internal or external command, operable program or batch file.' occurs, try the alternate installation
commands below.

ALTERNATE INSTALLATION COMMANDS:
python3 -m pip install pandas
python3 -m pip install requests
python3 -m pip install time
python3 -m pip install os

Next, double click on the [FOIA]WebCrawler.py file to execute the program. A Command Prompt Terminal will open and the user will be asked whether
they would like to update the Excel spreadsheet or not. Updating the Excel Spreadsheet inserts any new records from the gridView.xlsx file into the
[FOIA]FOIA Requests.xlsx file. Choosing to not update the file gives the user the option to webscrape details on any of the rows. The user can look
at the spreadsheet and then choose a row to start at and a row to end at. The script will find details for all rows in between the provided rows
(endpoints inclusive). 

NOTE: To fully update the sheet, download an updated gridView.xlsx file from the CPS public archive export button and
then replace the file in the folder of the python script. Then run the python script and type "y", to update the sheet.
Then run the program again and update rows accordingly.