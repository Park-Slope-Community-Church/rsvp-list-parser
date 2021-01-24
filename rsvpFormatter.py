import openpyxl
import constant
import csv
import collections
from pathlib import Path
from tkinter import Tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename

totalSeats = constant.capacity
class RSVP: 
    def __init__(self, first, last, age, service, primary):
        self.first = first
        self.last = last
        self.age = age
        self.service = service
        self.primary = primary

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
print(filename)

excelFile = filename
wbObj = openpyxl.load_workbook(excelFile)
wkSheet = wbObj.active

startDate = input("Enter a cutoff date in YYYY-MM-DD format: ")
rowStart = int(input("Enter the row number to start: "))
checkForNew = input("Point out new people? (yes/no)")

dependantExcludeList = ['Grade: ', 'Email address']
rsvpList1 = []
rsvpList2 = []
fullNameList = []
for row in wkSheet.iter_rows(min_row = rowStart):
    rowDate = row[4].value.split()[0]
    fullName = row[1].value + " " + row[2].value
    fullNameList.append(fullName.upper())
    if startDate <= rowDate:
        newRsvp = RSVP(row[1].value, row[2].value, "adult", row[8].value, row[2].value)
        if newRsvp.service == "10:00am":
            rsvpList1.append(newRsvp)
        elif newRsvp.service == "11:30am":
            rsvpList2.append(newRsvp)
        
        if row[7].value != None:
            dependants = row[7].value.split('\n')
            for dependant in dependants[:]:
                if dependant == '':
                    dependants.remove(dependant)
                else: 
                    for exclude in dependantExcludeList:
                        if dependant.startswith(exclude) == True:
                            dependants.remove(dependant)

            print(dependants)
            for dependant in dependants[:]:
                dependentComps = dependant.split(' ')
                dependentComps = [i for i in dependentComps if i != '']
                if len(dependentComps) > 0:
                    depRsvp = RSVP(dependentComps[0], dependentComps[1], dependentComps[2].replace('(', '').replace(')', ''), row[8].value, row[2].value)
                    if depRsvp.service == "10:00am":
                        rsvpList1.append(depRsvp)
                    elif depRsvp.service == "11:30am":
                        rsvpList2.append(depRsvp)

total1 = 0
total2 = 0
rsvpList1.sort(key = lambda rsvp: rsvp.primary)
rsvpList2.sort(key = lambda rsvp: rsvp.primary)
nameCounts = collections.Counter(fullNameList)
print(nameCounts)
for rsvp in rsvpList1:
    if rsvp.age == 'child':
        print("    " + rsvp.first + " " + rsvp.last + " (" + rsvp.age + ")")
    else:
        total1 = total1 + 1
        fullName = rsvp.first + " " + rsvp.last
        if nameCounts[fullName.upper()] == 1:
            print(fullName + " " + "--NEW--")
        else:
            print(fullName)

print("Total RSVP’s for worship: " + str(total1))
print("Remaining Seats: " + str(constant.capacity - total1))

print("CUT OF LINE ----------------")
for rsvp in rsvpList2:
    if rsvp.age == 'child':
        print("    " + rsvp.first + " " + rsvp.last + " (" + rsvp.age + ")")
    else:
        total2 = total2 + 1
        fullName = rsvp.first + " " + rsvp.last
        if nameCounts[fullName.upper()] == 1:
            print(fullName + " " + "--NEW--")
        else:
            print(fullName)
            
print("Total RSVP’s for worship: " + str(total2))
print("Remaining Seats: " + str(constant.capacity - total2))



#TODO - Remove duplicate names if present
#TODO - Craete Helper Methods
#TODO - Remove duplicated last names
#TODO - Bug where some kids are getting marked as adults