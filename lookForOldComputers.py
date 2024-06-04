#Imported Libraries

#openpyxl - Library for reading and manipulating Excel Spreadsheets
import openpyxl

#datetime - Library for working with dates
from datetime import date

#ldap3 - Library for connecting to AD
from ldap3 import Connection

#Function to prompt the user for desired settings
def prompt_user() :
    workbook_path = input("Enter the file path of the excel workbook containing the Blubeam Computer data >> ")
    print()
    last_ping_column_name = input("Enter the header name of the column that contains the Last Ping dates >> ")
    print()

    while True :
        try :
            max_days_old = int(input("Enter the maximum days since last ping >> "))
            print()
        except :
            print()
            print("Please enter an integer.")
            print()

    ip_address = input("Enter your ")
    
            
#Function to identify the old Blubeam computers from an Excel spreadsheet
def identify_old_computers(max_days_old, workbook_path, last_ping_column_name) :

    #Load Excel sheet 
    workbook = openpyxl.load_workbook(workbook_path)

    #Get the active sheet
    sheet = workbook.active 

    #Print message
    print("Here are the computers older than "+str(max_days_old)+" days old.")
    print()

    #Iterate through each column
    for column in sheet.iter_cols() :

        #If the header of the column is the Last Ping column
        if column[0].value == last_ping_column_name :

            #Iterate through each cell in the column
            for index, last_ping_cell in enumerate(column[1:]) :

                #Get the date of last ping as a string
                last_ping = last_ping_cell.value

                #Get the year, month, and day of the last ping in integer format
                year, month, day = [int(num) for num in last_ping.split("T")[0].split("-")]

                #Format the year, month, and day into a datetime date() object
                date_last_ping = date(year, month, day)

                #Get today's date in a date() object
                today = date.today()

                #Get the difference in days between today and the date since last ping
                difference = (today-date_last_ping).days

                #If the difference is greater than the maximum days old input
                if difference > max_days_old :

                    #Iterate through each cell in its row
                    for cell in sheet[index+2] :

                        #If the first three characters of the cell's value is GRE, it is the device name
                        if cell.value[:3] == "GRE" :

                            #Print the device name
                            print(cell.value)

