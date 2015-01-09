# ams-gradebook
A script to ease entering grades into Google sheets.

# Usage
This script makes entering grades a little easier.  Normally entering grades into Google sheets
requires alphabetizing the assignments by hand, and manually inputting the grades into a 
gradebook spreadsheet hosted on Google sheets.  The script provided takes a list of student's
first initials, last names, and scores and merges them into the Google sheet gradebook
using the [Google sheets API](https://developers.google.com/google-apps/spreadsheets/).

The script is written in Python2.  To use, run:
```
python gradebook.py grades_input_file.txt
```
where `grades_input_file.txt` is a file containing the grades to be merged in the spreadsheet.
The first line of `grades_input_file.txt` should be the assignment's header in the Google sheet.
Each line thereafter consists of the student's first initial, last name, and grade separated by
spaces.  An example grades file, example_grades.txt, is provided.

Before running the script, the gradebook's Google sheet's key must be set in spreadsheetkey.py.
Grades that could not be merged successfully into the gradebook are stored in unmergeable_students.txt
after the script has finished.

# Dependencies: 
[gspread](https://github.com/burnash/gspread)

[oauth2client](https://github.com/google/oauth2client)
