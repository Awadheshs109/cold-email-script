import smtplib
import os
import pandas as pd

from dotenv import load_dotenv
load_dotenv()
# Email Configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

def send_confirmation_email(email, company, name):
    """Sends confirmation email to the user."""
    subject = f"Regarding Internship/Job Opportunity at {company}"
    body = f"""
Your body of the mail
    """
    message = f"Subject: {subject}\n\n{body}"
    
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)  
            server.sendmail(EMAIL_SENDER, email, message) 
        print(f"Email sent to {company}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        print("There was an issue while sending the email to {company}, mail is {mail}.")
        
def run_app():
    data = pd.read_csv("Hr-detail.csv")
    emails = data['Email']
    company_name = data['Company']
    for index, i in data.iterrows():
        # print(i['Name'], i['Email'], i['Company'])
        send_confirmation_email(i['Email'], i['Company'], i['Name'])
    
run_app()