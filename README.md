# Auto_email_sender
#### This program automates sending emails to multiple recipients listed in an Excel sheet within seconds. The script reads the Excel file, checks the background color of rows, and sends emails only to those with green or orange backgrounds. It uses Python, Pandas, OpenPyXL, and smtplib for seamless email automation.
</br>

#### You can modify the script to fit your specific requirements, such as adjusting column indices for email addresses and names or filtering recipients based on different criteria.

## How To Run:
1) Download the dependencies pandas, openpyxl, python-dotenv.</br>
2) Set your sender's email id in line number 18. </br>
3) In line 13, set the sheet name in which the details are stored.</br>
4) Set your Google Apps 16 digit passcode in the .env file.</br>
5) On line 23 you can set the line from where the details of the participants start. (here 0 as there was no header row) </br> 
6) On line 24 and 25 you can set which column of the excel sheet is the email and name stored respectively.</br>
7) Line 35-48 contains the subject and body of the message which you can modify it as per your use.</br>
8) You can set the port number and smtp server as per requirements.</br>
9) Now you can run the script and emails will be sent. 
