# Declare arrays to store raw data from CSV Files
import csv
f=open("C:\Users\chious\Box Sync\vba-projects\pr-status\week0\v3-test-data-03-13-2019.csv")
csv_f=csv.reader(f)
for row in csv_f:
    print row
