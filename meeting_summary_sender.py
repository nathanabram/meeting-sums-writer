from __future__ import print_function

import gspread

import emailsender
import smtplib

from datetime import date

from oauth2client.service_account import ServiceAccountCredentials



#########STUFF FROM GOOGLE THING FOR GETTING THE CREDS FOR GMAIL ################
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

scope= ['https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive']

#credentials = ServiceAccountCredentials.from_json_keyfile_name('meeting-summaries-9d242890b8c3.json', scope)


gc = gspread.service_account()
sh = gc.open("Meeting Summaries")

summaries = sh.sheet1.get_all_values()
stuData = sh.worksheet("Student Database").get_all_values()


############################## SPREADSHEET NAV TOOLS ####################################


#Function to find the most recent entry

def last_filled_row(worksheet) :
    str_list = list(filter(None, worksheet.col_values(1)))
    return len(str_list)


#Return the column number of a specific lable, given a worksheet
def find_col_by_id(worksheet, name) :
    col_lable = worksheet[0].index(name)
    return col_lable

#Now working to find the last parsed entry.
#    Assumption: worksheet uses "parsed:" column with y/n/null values
def last_parsed_row(worksheet) :
    parsed_bool_column = worksheet[0].index("parsed:")
    cnt = 0
    #Find the number of rows that have been parsed
    for x in range(len(summaries)) :


        if (summaries[x+1][parsed_bool_column] =="y") :
            cnt=cnt +1
        else :
            return cnt + 1
            break

#Mark out all of the parsed rows as completed (on user request)
def MarkAllParsed(worksheet):
    markParsed = input("Mark these as parsed (y/n)")
    if (markParsed == "y"):
        for x in range(cellRange):
            sh.sheet1.update_cell(last_parsed_row(worksheet) + 1 +x,worksheet[0].index("parsed:")+1, "y")


#find cell with given row and column lable
def find_cell(worksheet, row, colName) :

    col = worksheet[0].index(colName) # Look in the first row of the worksheet and return the number that goes with the desired column name
    return worksheet[row][col] # Return cell contents



def find_rowcol_by_contents(worksheett, contents) :
	for i in range(len(worksheett)) :
		for j in range(len(worksheett[i])):
			if worksheett[i][j] == contents :
				return [i,j]
	else:
		return "null"



#Get day of the week from a given date

#Turn a m/d/yyyy date into a yyyy-mm-dd
#    so that date.fromisoformat('yyyy-mm-dd') can be used
#    Disregards anything after the year


def date_to_isoformat(rawDate) :
    mdSlash = rawDate.index("/", 0, 3)
    dySlash = rawDate.index("/", 3, 6)
    month = rawDate[:mdSlash]
    day = rawDate[mdSlash+1:dySlash]
    year = rawDate[dySlash + 1 : dySlash + 5]

    if len(month) == 1 :
        month = "0" + month
    if len(day) == 1 :
        day = "0" + day
    isoDate = year + "-" + month + "-" + day
    return str(isoDate)

##################### DATE STUFF #############################
#Given a day in iso format, return the weekday:
#
#
def iso_to_weekday(isoDate) :
    dateObj = date.fromisoformat(isoDate)
    weekdayIndex = dateObj.weekday()

    weekdays = {
        0 : "Monday",
        1 : "Tuesday",
        2 : "Wednesday",
        3 : "Thursday",
        4 : "Friday",
        5 : "Saturday",
        6 : "Sunday"
    }
    
    return weekdays[weekdayIndex]
#Take a Google sheets date cell value and return the weekday
def date_cell_to_weekday(dateValue) :
    return iso_to_weekday(date_to_isoformat(dateValue))
#
#
#


######################## CREATE LIST OF STUDENTS ###########################
#
#
#
# Based on this list, the final output is created, by looping and checking for these names

#Initialize the list of student names for this week's reports
stu_list = []
#Find which column has the student values
stu_name_col = summaries[0].index("Student (first and last):")

# Find the space of all the entries from this week
lastParsed = last_parsed_row(summaries)
cellRange = len(summaries) - lastParsed

#Create a non-repeating list of the students
for x in range(cellRange):
  #Take the name of the student in the next unchecked, unparsed row
    stuName = summaries[lastParsed  + x][find_col_by_id(summaries, "Student (first and last):")]
  #if its not a duplicate, add it to the list
    if (stuName not in stu_list):
        stu_list.append(stuName)

#########TODO: Tutor Payment Integration ###############
#tutorsToPay = [[]]

#for x in range(cellRange):
    #Next, figure how many hours each tutor worked.
#    tutorEmail = find_cell(summaries, lastParsed + 1 + x, "Email")
#    if tutorEmail not in tutorsToPay :
#        tutorsToPay.append([tutorEmail, ])


####################### WRITE SUMMARIES ###########################

#Create a list to hold the summaries for each student
sumsText = []
for x in range(len(stu_list)) :
    sumsText.append("")
    for y in range (0, cellRange) :# for each entry this week
        stuName = summaries[lastParsed + y][find_col_by_id(summaries, "Student (first and last):")]
        if (stuName == stu_list[x]) : #If the student in the entry on row y is the student whose sumsText is being created (the x), then
            weekdayLable = date_cell_to_weekday(find_cell(summaries, lastParsed  + y, "Timestamp"))
            notesLink = find_cell(summaries, lastParsed  + y, "Bitpaper link")
            summary = find_cell(summaries, lastParsed  + y, "Meeting summary (this will be included in the weekly summaries to the student and client)")
            summary = "\n" + weekdayLable + "\n" + summary + " \nNotes:\n" + notesLink 
            sumsText[x] = sumsText[x] + "\n" + summary
            

############################# ASSEMBLE AND SEND EMAILS ########################


testingStuff = []
for x in range(len(sumsText)) :
	messagetosend = "Here is %s's summary for this week:" % (str(stu_list[x].split(" ")[0])) + str(sumsText[x]) #Get the first name and insert, and add on the summary section
	databasespot = find_rowcol_by_contents(stuData, stu_list[x]) # Look in the second sheet which has the student emails to send to and so forth. This does return null right?
	if (databasespot != "null") :
		send_to = find_cell(stuData, databasespot[0], "Emails to Recieve Summaries")
		emailsender.send_email("MosaicMath Weekly Summary", messagetosend	, send_to)
		testingStuff.append([messagetosend, send_to])
	else:
		send_to = "nathan@mosaicmath.com"
		messagetosend = "***This did not go out to anyone. Please make sure that there's an email associated with this student in the Student Database***\n\n\n\n" + messagetosend
		emailsender.send_email("Error in weekly summary", messagetosend, send_to)

        
MarkAllParsed(summaries)
















