# Simon Alland
# Booking Engine
# This program is a simple hotel booking engine, which has the ability to create and look up reservations, as well as check in and check out guests. 
# Additionally, hotel configuration is stored in a config file which can be modified from a setup menu.
# Might need to install pyodbc in case Python is unable to import: pip install pyodbc
# Need to install AccessDatabaseEngine_X64.exe for this to work: https://www.microsoft.com/en-US/download/details.aspx?id=13255
# References: 
#   https://github.com/mkleehammer/pyodbc/wiki/Connecting-to-Microsoft-Access
#   https://www.geeksforgeeks.org/clear-screen-python/
#   https://stackoverflow.com/questions/17361338/convert-string-to-date-in-ms-access-query
#   https://www.geeksforgeeks.org/python-program-to-replace-specific-line-in-file/#

import re
import os
import pyodbc
from datetime import date, timedelta, datetime

def getConfiguration(configContent):
    # Create dictionary for rates and inventory
    inventory = dict()
    rates = dict()
    rooms = dict()
    roomRanges = dict()

    # Regex for hotelName and roomTypes
    regexHotel = "(?<=HotelName: ).*"
    regexHotelAbrv = "(?<=HotelAbbreviation: ).*"
    regexRoomType = "(?<=RoomType: ).*"
    
    # Parse through config file
    for line in configContent:
        hotelNames = re.findall(regexHotel,line)
        hotelAbrvs = re.findall(regexHotelAbrv,line)
        roomType = re.findall(regexRoomType,line)

        if hotelNames:
            hotelName = hotelNames[0]
        if hotelAbrvs:
            hotelAbrv = hotelAbrvs[0]
        if roomType:
            # Split row on the colon character
            roomType = roomType[0].split(":")
            
            # Append Room Type and Inventory to the inventory dictionary
            inventory[roomType[0]] = roomType[1]
            
            # Append Room Type and Rates to the Rates dictionary
            rates[roomType[0]] = roomType[2]
            
            # Get room numbers for roomtype
            roomRange = roomType[3].split('-')
            startRoom = int(roomRange[0])
            # Adding 1 to include last room
            endRoom = int(roomRange[1]) + 1
            
            # Loop through range of rooms and add to rooms dictionary
            for room in range(startRoom,endRoom):
                rooms[room] = roomType[0]
    
            # Adding the string ranges to dictionary for Setup menu
            roomRanges[roomType[0]] = roomType[3]

    # Return config values
    return hotelName, hotelAbrv, inventory, rates, roomRanges, rooms

def mainMenu(hotelName):
    # Set menuOption variable
    menuOption = ""

    # Keep menu running until a proper option is set
    while(menuOption == ""):
        print("\nBooking Tool Main Menu\n")
        print("Hotel: ",hotelName)
        print("1. New Booking")
        print("2. Look Up Booking")
        print("3. Check in Guest")
        print("4. Check out Guest")
        print("5. Hotel Setup")
        print("9: Quit\n")
        menuOption = input("What do you want to do?: ")
        # Test to see if a valid integer is given
        try: int(menuOption)
        except:
            menuOption = ""
            # Clear screen
            os.system('cls')
    return menuOption

def newBooking(hotelName, hotelAbrv, inventory, rates):
    # Create isAvailable variable
    isAvailalable = ''
    
    # Clear screen
    os.system('cls')
    
    while(isAvailalable != True):
        print("New Booking\n")
        print("Hotel: ",hotelName)
        
        # Enter in values for new booking
        guestName = input("Guest Name: ")
        roomType = input("Room Type: ")
        arrivalDate = input("Arrival Date: ")
        lengthStay = input("Length of Stay: ")
        
        try:
            for room in inventory:
                # Check to see if roomtype exists in inventory dictionary
                if roomType == room:
                    # Get rate associated with that roomtype
                    rate = rates[roomType]
        except: 
            print("Invalid Room chosen!")

        try: int(lengthStay)
        except: print("Length of stay needs to be a number!")

        # Check inventory on arrival date
        isAvailalable = checkInventory(arrivalDate, roomType, lengthStay, inventory)
        
        # Generate ID and save booking if inventory is available
        if isAvailalable == True:
            # Generate reservation ID
            resID = newResID(hotelAbrv)
            # Save Booking info
            addBooking(hotelAbrv, resID, guestName, roomType, rate, arrivalDate, lengthStay)
            print("Reservation ID is: ",resID)
        
        # Display no availability if not available
        elif isAvailalable == False:
            print("No available rooms!")

def checkInventory(arrivalDate, roomType, lengthStay, inventory):
    # Create availability dictionary
    availability = dict()

    # Create a matrix of dates and room types with available rooms
    query = "select ArrivalDate, LengthStay, RoomType from bookings"
    results = databaseQueryAllResults(query)

    # Loop through each booking in database
    for item in results:
        
        # Go through all multi-day bookings
        for day in range(0, item[1]):
            dateStay = item[0] + timedelta(days = day)
            
            # Make tuples of date and roomtype
            if (dateStay,item[2]) in availability:
                availability[dateStay,item[2]] += 1
            else:
                availability[dateStay,item[2]] = 1

    # Go through reqeuested booking dates
    for day in range(0, int(lengthStay)):
        dateStay = datetime.strptime(arrivalDate, '%m/%d/%Y') + timedelta(days = day)        
        
        # If there is no tuple of this date, skip to next iteration
        if (dateStay,roomType) not in availability: continue      
        
        # If requested date has no availability return False
        if availability[dateStay,roomType] >= int(inventory[roomType]):
            return False
    
    # Return true if not failing above test
    return True

def newResID(hotelAbrv):
    # Create query for reservation ID
    query = "select count(*) from bookings where HotelCode = '" + hotelAbrv +"'"
    
    # Connect to database and run query
    resCount = databaseQuery(query)

    # Create reservation ID - format is PSM-1, etc
    resCountID = resCount[0]+1
    resID = hotelAbrv + "-" + str(resCountID)

    return resID

def databaseQuery(query):
    # Create database connection string and connect to database
    connectionString = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=.\Final Project\bookings.accdb;'
        )
    dbConnection = pyodbc.connect(connectionString)
    dbCursor = dbConnection.cursor()    

    # Execute query
    dbCursor.execute(query)
    result = dbCursor.fetchone()

    # Close connection
    dbConnection.close()

    return result

def databaseQueryAllResults(query):
    # Create database connection string and connect to database
    connectionString = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=.\Final Project\bookings.accdb;'
        )
    dbConnection = pyodbc.connect(connectionString)
    dbCursor = dbConnection.cursor()

    # Execute query
    dbCursor.execute(query)
    result = dbCursor.fetchall()

    # Close connection
    dbConnection.close()

    return result

def databaseUpdate(query):
    # Create database connection string and connect to database
    connectionString = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=.\Final Project\bookings.accdb;'
        )
    dbConnection = pyodbc.connect(connectionString)
    dbCursor = dbConnection.cursor()    

    # Execute query
    dbCursor.execute(query)
    
    # Commit Query
    dbConnection.commit()

    # Close connection
    dbConnection.close()    

def addBooking(hotelAbrv, resID, guestName, roomType, rate, arrivalDate, lengthStay):
    # Create query to add booking
    query = "insert into bookings (HotelCode, ResID, GuestName, RoomType, Rate, ArrivalDate, LengthStay) values ('" + hotelAbrv + "','" + resID + "','" + guestName + "','" + roomType + "','" + rate + "','" + arrivalDate + "'," + lengthStay + ")"
    
    # Execute query
    databaseUpdate(query)

def getBooking(hotelName):
    # Clear screen
    os.system('cls')
    print("Search for Booking\n")
    print("Hotel: ",hotelName)
        
    # Enter in values to look up booking
    resID = input("Reservation ID: ")

    # Create query to search for booking
    query = "select * from bookings where ResId = '" + resID + "'"
    result = databaseQuery(query)
    
    # Output results
    if result:
        print("\nGuest Name: ",result[2])
        print("Room Type: ",result[3])
        print("Room Rate: ",result[4])
        print("Arrival Date: ",result[5].strftime('%m/%d/%Y'))
        print("Length of Stay: ",result[6])
    else: print("\nReservation Not Found!")

def checkin(hotelName, rooms):
    # Clear screen
    os.system('cls')
    print("Check in Guest\n")
    print("Hotel: ",hotelName)
    
    # Enter in values to look up booking
    resID = input("Reservation ID: ")
    
    # Check to see if reservation exists
    query = "select count(*) from bookings where ResId = '" + resID + "'"
    result = databaseQuery(query)

    # Check in guest if result found
    if result[0] == 1:
        
        # Check to see if reservation is already checked in
        query = "select count(*) from bookings where ResId = '" + resID + "' and InHouse = 0"
        result = databaseQuery(query)
        
        if result[0] == 1:
            # Get todays date
            todaysDate = date.today()
            # Check to see if reservation is arriving today         
            query = "select count(*) from bookings where ResId = '" + resID + "' and InHouse = 0 and ArrivalDate = DateValue('" + str(todaysDate) +"')"
            result = databaseQuery(query)

            if result[0] == 1:
                # Get room number on checkin    
                roomNumber = getRoomNumber(resID, rooms)

                # Create query to check in guest
                query = "Update bookings set inHouse = 1 where ResId = '" + resID + "'"
                databaseUpdate(query)

                # Output success message
                print("\nGuest Checked in Successfully!")
                print("Room number assigned: ",roomNumber)
            
            # Print error if arrival date is not today
            else: print("Arrival date is not today!")
            
        # Print error if guest is already checked in
        else: print("Reservation is already checked in!")
        
    # Print error if reservation not found
    else: print("Reservation not found!")

def getRoomNumber(resID, rooms):    
    # Create query to get arrival date and other data from booking
    query = "select ArrivalDate, RoomType, LengthStay from bookings where ResId = '" + resID + "'"
    result = databaseQuery(query)
    arrivalDate = str(result[0])
    roomType = result[1]
    lengthStay = result[2]

    # Loop through rooms dictionary
    for room in rooms:
        # Verify we are only looking at matching roomtype
        if roomType == rooms[room]:
            # Run function to check if room is available
            isAvailable = checkRoomAvailable(arrivalDate, room, lengthStay)
            # if true, assign room and return
            if isAvailable == True:
                # Update reservation with room number assigned
                query = "update bookings set RoomNumber = " + str(room) + " where ResId = '" + resID + "'"
                databaseUpdate(query)
                return room

def checkRoomAvailable(arrivalDate, roomNumber, lengthStay):
    # Create availabilityRooms dictionary
    availabilityRooms = dict()

    # Create a matrix of dates and room types with available rooms
    query = "select ArrivalDate, LengthStay, RoomNumber from bookings"
    results = databaseQueryAllResults(query)

    # Loop through each booking in database
    for item in results:
        
        # Go through all multi-day bookings
        for day in range(0, item[1]):
            dateStay = item[0] + timedelta(days = day)
            
            # Make tuples of date and room number
            if (dateStay,item[2]) in availabilityRooms:
                # This should never be greater than 1
                availabilityRooms[dateStay,item[2]] += 1
            else:
                availabilityRooms[dateStay,item[2]] = 1

    # Go through reqeuested booking dates
    for day in range(0, int(lengthStay)):
        dateStay = datetime.strptime(arrivalDate,'%Y-%m-%d %H:%M:%S') + timedelta(days = day)        
        
        # If there is no tuple of this date, skip to next iteration
        if (dateStay,roomNumber) not in availabilityRooms: continue      
        
        # If requested date has no availability return False
        if availabilityRooms[dateStay,roomNumber] >= 1:
            return False
    
    # Return true if not failing above test
    return True                

def checkout(hotelName):
    # Clear screen
    os.system('cls')
    print("Check Out Guest\n")
    print("Hotel: ",hotelName)
    
    # Enter in room number to look up booking
    roomNumber = input("Room Number: ")
    
    # Check to see if reservation is already checked in and get resID
    query = "select resId from bookings where RoomNumber = " + roomNumber + " and InHouse = 1"
    result = databaseQuery(query)

    if result: 
        # Set resId
        resID = result[0]

        # Create query to check out guest
        query = "Update bookings set inHouse = 0 where ResId = '" + resID + "'"
        databaseUpdate(query)

        # Get final bill
        query = "select Rate, LengthStay from bookings where ResId = '" + resID + "'"
        result = databaseQuery(query)
        finalBill = format((int(result[0]) * int(result[1])),".2f")

        # Output success message
        print("\nGuest Checked Out Successfully!")
        print("Final bill is: $",finalBill,"")
    
    # Print error if reservation is checked out
    else: print("Reservation is already checked out!")


def setup(configFile, hotelName, inventory, rates, roomRange):
    # Clear screen
    os.system('cls')
    
    # Set menuOption variable
    menuOption = ""
    
    # Keep menu running until a proper option is set
    while(menuOption == ""):
        print("\nBooking Tool Setup Menu\n")
        print("1. Add/Edit Roomtypes")
        print("2. Delete Reservation")
        print("9. Return to main menu\n")
        menuOption = input("What do you want to do? ")
        # Test to see if a valid integer is given
        try: int(menuOption)
        except:
            menuOption = ""
            # Clear screen
            os.system('cls')
        if menuOption == '1': setRoomTypes(configFile, hotelName, inventory, rates, roomRange)
        elif menuOption == '2': deleteReservation(hotelName)
        # Quit to main menu
        elif menuOption == '9': return

def setRoomTypes(configFile, hotelName, inventory, rates, roomRange):
    # Clear screen
    os.system('cls')
    
    print("Add/Edit Roomtypes\n")
    print("Hotel: ",hotelName,"\n")
    # Enter the abbreviation of a roomtype
    roomType = input("Please enter in a roomtype to add/edit: ")
    
    # Check if roomtype exists in current config
    if roomType in inventory:
        print("Roomtype exists!\n")
        # Load in values from roomtype
        print("Roomtype: ",roomType)
        print("Inventory: ",inventory[roomType]," rooms.")
        print("Rate: $",format(int(rates[roomType]),".2f"))
        print("Room Numbers",roomRange[roomType])

    # Enter in values to add/edit room
    print("*** Edit Mode ***")
    newRoomType = input("Roomtype: ")
    newInventory = input("Total Inventory: ")
    newRate = input("Room Rate: ")
    newRoomNumbers = input("Room Numbers (#-#): ")

    # Check if roomtype given is in inventory
    if newRoomType not in inventory:
        # Create config string
        roomConfig = "RoomType: " + newRoomType + ":" + newInventory + ":" + newRate + ":" + newRoomNumbers +"\n"

        # Open config in append mode
        configContent = open(configFile,'a')
        configContent.write(roomConfig)
        configContent.close

    if newRoomType in inventory:
        # Set list counter
        counter = 0

        roomConfig = "RoomType: " + newRoomType + ":" + newInventory + ":" + newRate + ":" + newRoomNumbers + "\n"

        # Open config in read mode and store lines in variable
        with open(configFile,'r') as file:
            data = file.readlines()

        # Loop through file looking for roomtype to replace
        for line in data:
            # Check if roomtype in this line
            if newRoomType in line:
                # Update config
                data[counter] = roomConfig
            
            # Increment list counter
            counter += 1

        # Write the changes to the file
        with open(configFile, 'w') as file:
            file.writelines(data)

def deleteReservation(hotelName):
    # Clear screen
    os.system('cls')
    print("Delete Reservation\n")
    print("Hotel: ",hotelName)
    
    # Enter in resID to look up booking
    resID = input("Reservation ID: ")
    # Check to see if reservation is already checked in and get resID
    query = "select resId from bookings where resID = '" + resID + "'"
    result = databaseQuery(query)

    if result: 
        # Create query to delete reservation
        query = "Delete from bookings where ResId = '" + resID + "'"
        databaseUpdate(query)
        print("Reservation Deleted!")
    else: print("Reservation not found!")

def main():
    # Set variables
    menuOption = ""

    # Clear screen
    os.system('cls')

    # Set config file name and open file
    configFile = "./Final Project/config.txt"
    configContent = open(configFile)
    
    # Get configuration information
    hotelName, hotelAbrv, inventory, rates, roomRange, rooms = getConfiguration(configContent)
    configContent.close

    while menuOption == "":
        # Go into main menu
        menuOption = mainMenu(hotelName)
        if menuOption == '1': newBooking(hotelName, hotelAbrv, inventory, rates)
        elif menuOption == '2': getBooking(hotelName)
        elif menuOption == '3': checkin(hotelName, rooms)
        elif menuOption == '4': checkout(hotelName)
        elif menuOption == '5': setup(configFile, hotelName, inventory, rates, roomRange)
        # Quit if 9 is given
        elif menuOption == '9': exit()
        menuOption = ""

main()
