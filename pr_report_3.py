# Declare arrays to store raw data from CSV Files
WeekNumber=input("Please Enter Week Number: ")
CutOff=input("Please enter cutoff date in the format of mm/dd/yyyy: ")
import csv
with open(r'''C:\Users\chious\Box Sync\vba-projects\pr-status\week0\v3-test-data-03-13-2019.csv''') as csvfile:
    readCSV=csv.reader(csvfile, delimiter=",")
    rawdata=[]
    headers=next(readCSV)
    for row in readCSV:
        rawdata.append(row)
