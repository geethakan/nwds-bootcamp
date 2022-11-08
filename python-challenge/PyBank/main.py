#Import dependent modules
import os
import csv
import sys
import math
import statistics as st

#declare global variables
net_profitloss = 0.0
net_change = 0.0
last_profitloss = 0.0
total_months = 0
chg_list = list()     #this challenge is done using list and not dictionary
chg_month = list()

#set input and output file names with path
in_filewpath = os.path.join("Resources", "budget_data.csv")
out_filewpath = os.path.join("analysis", "financial_report.txt")

#open input csv and read
with open(in_filewpath, "r") as ifile:
    reader = csv.reader(ifile, delimiter=",")
    header = next(reader)

    #with header skipped, work with data rows in csv
    for row in reader:

        #total count of rows and profit/loss (column 2)
        total_months += 1
        net_profitloss += float(row[1])

        #add profit/loss change value to list - need to load from 2nd row; also store chg month
        if total_months > 1:            #Note: this is not index but row count
            chgvalue = float(row[1]) - last_profitloss
            chg_list.append(chgvalue)
            chg_month.append(row[0])

        #save profit/loss value to use in next row comparison
        last_profitloss = float(row[1])

#Get min/max value index Note: getting just first index if min/max is in multiple rows
min_index = chg_list.index(min(chg_list))
max_index = chg_list.index(max(chg_list))

# save a reference to the original standard output
original_stdout = sys.stdout

with open(out_filewpath, "w") as ofile:
    sys.stdout = ofile                  #changing standard outout to output file

    # print count of months and net profit/loss
    #to format results, math.floor used to get integer portion; .2f used for 2 decimal places
    print("Financial Analysis")
    print("---------------------------------------------------")
    print(f"Total Months: {total_months}")
    print(f"Total: ${math.floor(net_profitloss)}")

    #use list functions to get sum, min and max of the change in profit/loss
    print(f"Average change: ${((sum(chg_list))/len(chg_list)):.2f}")
    print(f"Greatest Increase in Profits: {chg_month[max_index]}  (${math.floor(max(chg_list))})")
    print(f"Greatest Decrease in Profits: {chg_month[min_index]}  (${math.floor(min(chg_list))})")

    sys.stdout = original_stdout        #Reset to original value

#Same print statements can be repeated here - but trying to open file and print values
with open(out_filewpath, "r") as ifile:
    for line in ifile:
        print(line.rstrip())