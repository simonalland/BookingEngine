# BookingEngine

Description:
    This program is a simple hotel booking engine, which has the ability to create and look up reservations, as well as check in and check out guests. Additionally, hotel configuration is stored in a config file which can be modified from a setup menu.

Prerequisites:
    Will likely need to install pyodbc in case Python is unable to import the library: pip install pyodbc
    Need to install AccessDatabaseEngine_X64.exe for this to work: https://www.microsoft.com/en-US/download/details.aspx?id=13255 
    I have the config file in the following location .\Final Project\config.txt -- make sure that structure exists.

Instructions:
    Run the program and choose what menu options to follow.
    1. Create a New Booking
    2. Look up an existing booking
    3. Check in a guest
    4. Check out a guest
    5. Hotel Setup
    9. Quit the program

Future Work:
    I would like to have better error checking for each inputted field, as well as having the ability to use multiple hotels for the program and being able to switch between them.
    Additionally it would be nice to include function rooms, but the idea for this would be for small bed and breakfasts to be able to book rooms for their guests.

References:
    https://github.com/mkleehammer/pyodbc/wiki/Connecting-to-Microsoft-Access
    https://www.geeksforgeeks.org/clear-screen-python/
    https://stackoverflow.com/questions/17361338/convert-string-to-date-in-ms-access-query
    https://www.geeksforgeeks.org/python-program-to-replace-specific-line-in-file/#
