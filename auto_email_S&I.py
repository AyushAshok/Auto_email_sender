import os
from email.message import EmailMessage
import pandas as pd
import openpyxl
import ssl
import smtplib
from dotenv import load_dotenv

load_dotenv()

file_path = "final_form.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb["Form Responses 1"]

GREEN_COLOR = "FF00FF00"
ORANGE_COLOR = "FFFF9900"

email_sender = 'aayush.ashok04@gmail.com'
password = os.environ.get("EMAIL_PASSWORD")

selected_students = []

for row in ws.iter_rows(min_row=0):
    email_cell = row[0]  # column is email
    name_cell = row[1]   # column is name
    
    # Check if cell has fill color
    if email_cell.fill.start_color.index in {GREEN_COLOR, ORANGE_COLOR}:
        email = email_cell.value
        name = name_cell.value if name_cell.value else "Student"
        if email:  # Only add if email exists
            selected_students.append({"email": email, "name": name})


subject = "Congratulations on your selection to S&I domain for AT'25"
body_template = """Dear {name},

We are thrilled to inform you that you have been selected for Stage And Infra domain for AT'25!

We were highly impressed by your skills, creativity, and enthusiasm, and we firmly believe that you will be a valuable addition to the team.

AT'25 is more than just an event—it's a journey filled with learning, collaboration, and unforgettable experiences. We are incredibly excited to have you onboard.

Please find the link for the WhatsApp group below:
WhatsApp Group Link - [INSERT_LINK_HERE]

Best regards,
AT'25 S&I Team"""

context = ssl.create_default_context()

try:
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(email_sender, password)
        i=0
        
        for student in selected_students:
            em = EmailMessage()
            em["From"] = email_sender
            em["To"] = student["email"]
            em["Subject"] = subject
            em.set_content(body_template.format(name=student["name"]))
            
            # smtp.sendmail(email_sender, student["email"], em.as_string())
            print(f"{i})Email sent to {student['name']} ({student['email']})")
            i+=1
            
    print(f"\nTotal emails sent successfully: {len(selected_students)}")
    
except Exception as e:
    print(f"An error occurred: {str(e)}")