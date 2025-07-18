# Clear the terminal screen (Windows only)
import os
os.system('cls')

# Import required libraries
import smtplib as s                              # For sending emails via SMTP
from email.message import EmailMessage           # For constructing the email content
import random                                    # For choosing a random birthday message
import pandas as p                               # For reading the Excel file
import datetime as d                             # For getting the current date and time

# -------------------------------
# Function to get current date and time
# -------------------------------
def DateTime():
    print("1")  # Debug print
    now = d.datetime.now()                      # Get current datetime
    current_date = now.day                     # Extract day
    current_month = now.month                  # Extract month
    current_hour = now.hour                    # Extract hour (not used here, but returned)
    return (current_date, current_month, current_hour)

# -------------------------------
# Function to read random birthday message
# -------------------------------
def Birthday_Message():
    print("3")  # Debug print

    # Read all non-empty lines from the message file
    with open("D:\\Programming\\programming\\python\\PYTHON BASIC\\Python_Learning\\Birthday Wisher (Day 32) start\\Birthday_wisher\\message1.txt", 
              "r", encoding="utf-8") as file:
        lines = [line.strip() for line in file if line.strip()]  # Strip empty lines

    random_wish = random.choice(lines)  # Choose a random message
    return random_wish

# -------------------------------
# Function to send birthday wishes via email
# -------------------------------
def Email_sender(data, now):
    print("2")  # Debug print
    recipients = []  # List to store (name, email) of today's birthdays

    # Loop through each row of the Excel data
    for i, r in data.iterrows():
        if now[0] == r["day"] and now[1] == r["month"]:  # If day and month match today's date
            person = r["Name"]
            email = r["email"]
            recipients.append((person, email))  # Add to recipients list

    # Email credentials (replace with your actual email and password or use environment variables)
    my_email = "rachitmishraagent47@gmail.com"                      
    password = "czvb vozc lpxp zfxj"

    # Loop through all matched recipients and send emails
    for person, email in recipients:
        random_wish = Birthday_Message()  # Get a random message

        # Construct the email
        msg = EmailMessage()
        msg["Subject"] = "Happy Birthday to YOU 🥳🎉🎉🎉"
        msg["From"] = my_email
        msg["To"] = email

        # Plain text content
        msg.set_content(f"Hey, {person}\n\n{random_wish}\n\nCheers!!\nFrom Rachit")

        # HTML version of the message
        msg.add_alternative(f"""\
        <html>
          <body>
            <p>Hey, {person}</p>
            <p><b>{random_wish}</b></p>
            <p>Cheers!!<br>From Rachit</p>
          </body>
        </html>
        """, subtype='html')

        # Connect to Gmail SMTP server and send the email
        with s.SMTP("smtp.gmail.com", 587) as con:
            con.starttls()                        # Start TLS encryption
            con.login(my_email, password)         # Login to Gmail
            con.send_message(msg)                 # Send the email

# -------------------------------
# Main Execution
# -------------------------------

# Load birthday data from Excel file
data = p.read_excel("D:\\Programming\\programming\\python\\PYTHON BASIC\\Python_Learning\\Birthday Wisher (Day 32) start\\Birthday_wisher\\Stored_Dates.xlsx")

# Call function to send birthday emails
Email_sender(data, DateTime())































# from datetime import datetime

# # Get the current datetime
# now = datetime.now()

# # Extract individual components
# current_date = now.day
# current_month = now.month
# current_year = now.year
# current_hour = now.hour
# current_minute = now.minute
# current_second = now.second

# # Print them
# print("Date:", current_date)
# print("Month:", current_month)
# print("Year:", current_year)
# print("Hour:", current_hour)
# print("Minute:", current_minute)
# print("Second:", current_second)











# czvb vozc lpxp zfxj