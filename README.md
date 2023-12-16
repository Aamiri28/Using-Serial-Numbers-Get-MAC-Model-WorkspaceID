# Using-Serial-Numbers-Get-MAC-Model-WorkspaceID
Code Explanation:

Purpose:

The script aims to automate interactions with a web application, extract information based on serial numbers from an Excel file, input these serial numbers into the web application, extract data from the web page, and save the extracted information into another Excel file.

Libraries Used:

selenium: Automates web browser interactions.
openpyxl: Handles Excel files in Python.


Code Breakdown:

Imports: Import necessary libraries/modules for web automation and Excel handling.

Web Automation Setup:
Initialize a Chrome WebDriver session.
Navigate to the specified URL.

Log In and Web Interactions:
Locate elements using XPath and interact with the webpage by entering login credentials, clicking buttons, and scrolling.
Wait for elements to load using WebDriverWait.

Excel Handling:
Load the Excel file containing serial numbers (serials.xlsx).
Create/open an output Excel file (output.xlsx) to store extracted information.
Set headers for the output Excel file.

Main Loop for Serial Numbers:
Loop through each row in the source Excel file.
Input the serial number into the web page's input field.
Extract information from the webpage based on the entered serial number.
Handle exceptions when information is missing.
Write extracted information (serial number, model, MAC, folder, customer ID) into the output Excel file.

Considerations:

Web Interactions: The script uses XPath locators to interact with elements on the web page. Ensure these locators are up-to-date and match the structure of the webpage.

WebDriver Compatibility: Ensure the installed Chrome WebDriver (chromedriver.exe) matches the installed Chrome browser version.

Exception Handling: The code includes exception handling to manage scenarios where expected information might not be found on the webpage.
