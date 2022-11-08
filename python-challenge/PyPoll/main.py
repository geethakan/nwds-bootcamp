#import dependent modules
import os
import csv

#declare global variables; dictionary for representatives and total votes
rep_dict = {}
total_votes = 0
winner = ""
winner_vote = 0

#set input file name with path
file = os.path.join("Resources", "election_data.csv")
out_txt = os.path.join("analysis", "election_results.txt")

#open csv in read mode
with open(file, "r") as f:
    reader = csv.reader(f, delimiter=",")
    header = next(reader)

    #we now have just data to loop thru (add to total vote count)
    for row in reader:
        k = row[2]
        total_votes += 1

        #if name in dictionary, add 1 to value; else add name as key and 1 as value
        if k in rep_dict:
            rep_dict[k] += 1
        else:
            rep_dict[k] = 1


#print results in terminal; also direct output to text file
with open(out_txt, "w") as f:
    print("Election Results")
    print("Election Results", file=f)
        
    print("------------------------------------------")
    print("------------------------------------------", file=f)

    print(f"Total Votes: {total_votes}")
    print(f"Total Votes: {total_votes}", file=f)

    print("------------------------------------------")
    print("------------------------------------------", file=f)

    #loop thru dictionary and print candidates name, percentage of votes(calculated value) and vote count
    #get name of winner while looping based on highest vote count (will be used to print later)
    for k, v in rep_dict.items():
        if v > winner_vote:
            winner_vote = v
            winner = k
        
        per_votes = (v/total_votes) * 100

        print(f"{k}: {per_votes:.3f}% ({v})")
        print(f"{k}: {per_votes:.3f}% ({v})", file=f)

    print("------------------------------------------")
    print("------------------------------------------", file=f)

    print(f"Winner: {winner}")
    print(f"Winner: {winner}", file=f)

    print("------------------------------------------")
    print("------------------------------------------", file=f)





