import csv

file_path = input("Enter file path:")

if not file_path.endswith(".txt"):
    print("Only .txt files are allowed"); exit()

try:
    curr_file = open(file_path, encoding="utf-8", newline="\n").readlines()
except FileNotFoundError as Err: 
    print(Err); exit()

hint = max(len(l.split()) for l in curr_file) #--> hint the user on how many columns is needed
field = input(f"Field columns separated by commas, must be of length {hint}:").split(",")

field_length = len(field)
if (not field) or (field_length < hint):
    emp_col = hint - field_length #--> get the max number of columns left to fill

    for fill in range(emp_col):
        field.append(f"None{fill}") #--> fill empty columns with None

with open("curr_file.csv", "w", encoding="utf-8") as file:
    csv_file = csv.DictWriter(file, field)
    csv_file.writeheader() #--> write header in the first row

    for lines in curr_file:
        line = lines.split()

        contact = {}
        n = 0 #--> make column (None + n) if no more columns

        for col in range(len(field)):
            try:
                contact[field[col]] = line[col]
            except IndexError:
                contact[field[col]] = None

        csv_file.writerow(contact) #--> write each line under its own column
