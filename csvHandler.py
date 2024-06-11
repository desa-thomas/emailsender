import csv

def readTable():
    tables = []
    with open('projectcmo.csv', 'r') as csvfile: 
        csv_reader = csv.reader(csvfile)
        header = next(csv_reader)
        currentCompany = ""

        for row in csv_reader:
            company = row[0]

            if company != currentCompany:
                print()

    return tables

print(readTable()) 
