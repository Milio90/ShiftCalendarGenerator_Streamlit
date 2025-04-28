# Purpose
I created this program in order to create .ics files from the files of medical personnel shifts in my centre.
# Requirements
The program was made with Python 3.12.3
The python packages used for this project are:
- python-docx
- icalendar
- streamlit

# How does it work
The program expects a .docx file with two tables. The first table should have the following format:

![image](https://github.com/user-attachments/assets/3822f819-242a-46c5-8531-6d711abef091)

Asterisks indicate on-call 24-hour shifts with the rest being regular 24-hour shifts

The second table should have this format:

![image](https://github.com/user-attachments/assets/2a142d04-6cee-4ebc-a5c8-a32034d80e26)

Optionally, the user can upload specialist on-call tables for the cath-lab and electrophysiology departments. These tables should be in two separate files with this format:

![image](https://github.com/user-attachments/assets/deeb6a3a-352f-42b6-a20d-3474dfeccc95)

Then, it creates .ics files with calendar events for the requested personnel's shifts. In the description of each event, it also adds the names of the other co-workers for the day.

The program tries to find the month and year automatically from the file name if given as "ΕΦΗΜΕΡΙΕΣ MONTH YEAR.docx" as well as from the contents of the tables. Otherwise, the user can specify them manually.

Unfortunately, it only works in Greek so far and only for docx files.
