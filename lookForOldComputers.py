#Imported Libraries

#openpyxl - Library for reading and manipulating Excel Spreadsheets
import openpyxl

#datetime - Library for working with dates
from datetime import date

#pyad - Library for connecting to AD
from pyad import *

#Function to prompt the user for desired settings
def prompt_user() :

    #Prompt user for the Excel file path
    workbook_path = input("Enter the file path of the excel workbook containing the Blubeam Computer data >> ")
    print()

    #Prompt user for the last ping column name in the Excel file
    last_ping_column_name = input("Enter the header name of the column that contains the Last Ping dates >> ")
    print()

    #Loop until user inputs valid integer
    while True :

        #Try to convert input into integer
        try :

            #Prompt user for the maximum amount of days old since last ping
            max_days_old = int(input("Enter the maximum days since last ping >> "))
            print()

            #Break loop if no exception
            break

        #If exception, print error message
        except :
            print()
            print("Please enter an integer.")
            print()
    
    #Prompt user for the name of the LDAP server
    ldap_server = input("Enter the name of your ldap server >> ")
    print()

    #Prompt user for username
    admin_user = input("Enter your admin username >> ")
    print()

    #Prompt user for password
    admin_passwd = input("Enter your admin password >> ")
    print()

    #Call "identify_old_computers"
    identify_old_computers(max_days_old, workbook_path, last_ping_column_name, ldap_server, admin_user, admin_passwd)
    
    
            
#Function to identify the old Blubeam computers from an Excel spreadsheet
def identify_old_computers(max_days_old, workbook_path, last_ping_column_name, ldap_server, admin_user, admin_passwd) :

    #Load Excel sheet 
    workbook = openpyxl.load_workbook(workbook_path)

    #Get the active sheet
    sheet = workbook.active

    #Connect to AD with pyad
    pyad.set_defaults(ldap_server= ldap_server, username=admin_user, password=admin_passwd)

    #Print message
    print("Here are the computers older than "+str(max_days_old)+" days old that are not in AD: ")
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

                            #Search for device name in AD
                            try :
                                computer_search = pyad.adcomputer.ADComputer.from_cn(cell.value)

                            #If not in AD, print device name
                            except :
                                
                                #Print the device name
                                print(cell.value)

prompt_user()
