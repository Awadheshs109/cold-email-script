import smtplib
import os
import pandas as pd
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

load_dotenv()

# Email Configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
RESUME_LINK = "https://drive.google.com/file/d/17fpc_EwLpJCKKFGPfzZiHuOamPHaILri/view?usp=sharing"  # Replace with your actual link

# List to track failed emails
failed_log = []

def send_confirmation_email(email, company, name):
    """Sends confirmation email with resume link in body."""
    subject = f"Regarding Job Opportunity at {company}"

    body = f"""\
Hi {name},<br><br>

I hope you're doing well.
I’m reaching out to express my interest in a <strong>Frontend Developer</strong> position at <strong>{company}</strong>.<br>
With over 3 years of experience in Angular, I’ve built scalable web applications using:
<ul>
  <li><strong>Angular (v13+)</strong> with standalone components</li>
  <li><strong>Reactive Forms</strong> for dynamic form control</li>
  <li><strong>NgRx</strong> for advanced state management</li>
  <li><strong>DevExtreme</strong>, <strong>Highcharts</strong>, and <strong>Power BI</strong> for dynamic dashboards</li>
</ul>


<a href="{RESUME_LINK}" style="display:none;" target="_blank">View My Resume</a>
<p>I’ve attached my resume and would welcome the opportunity to discuss how my technical skills and problem-solving mindset can contribute to your team.</p>

<p style="margin-bottom: 3px;">Regards,</p>

<strong style="font-size: 14px; color: #2980b9;">Awadhesh Sharma</strong><br>
<span style="font-size: 12px; color: #666;">Frontend Developer | Angular</span><br>
<span style="font-size: 12px; color: #666;">📞 +91 7571869952</span><br>

<div style="margin-top: 6px;">
  <a href="https://wa.me/917571869952" target="_blank" style="margin-right: 6px;">
    <img src="https://cdn2.iconfinder.com/data/icons/social-media-2285/512/1_Whatsapp2_colored_svg-512.png" width="16">
  </a>
  <a href="https://www.linkedin.com/in/awadheshs109" target="_blank" style="margin-right: 6px;">
    <img src="https://cdn3.iconfinder.com/data/icons/social-media-2068/64/_LinkedIn-64.png" width="16">
  </a>
  <a href="https://github.com/awadheshs109" target="_blank" style="margin-right: 6px;">
    <img src="https://cdn3.iconfinder.com/data/icons/social-media-2068/64/_github-64.png" width="16">
  </a>
  <a href="https://awadhesh-portfolio.vercel.app/" target="_blank">
    <img src="https://awadhesh-portfolio.vercel.app/favicon.ico" width="16">
  </a>
</div>
"""


    try:
        msg = MIMEMultipart()
        msg["From"] = EMAIL_SENDER
        msg["To"] = email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "html"))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)

        print(f"✅ Email sent to {name} at {company} ({email})")

    except Exception as e:
        print(f"❌ Failed to send email to {name} at {company} ({email})")
        print(f"   Error: {e}")
        failed_log.append({
            "Name": name,
            "Email": email,
            "Company": company,
            "Error": str(e)
        })

def run_app():
    try:
        # Read CSV with unknown extra columns
        df = pd.read_csv("Hr-detail.csv", header=0, engine='python')
    except FileNotFoundError:
        print("❌ 'Hr-detail.csv' not found.")
        return

    for index, row in df.iterrows():
        name = str(row.get("Name", "")).strip()
        email = str(row.get("Email", "")).strip()
        company = str(row.get("Company", "")).strip()

        if not email or not name or not company or "@" not in email:
            print(f"⚠️ Skipping incomplete row {index + 2}")
            continue

        send_confirmation_email(email, company, name)

    if failed_log:
        print(f"\n❗ {len(failed_log)} email(s) failed. Logging to failed_emails.csv")
        pd.DataFrame(failed_log).to_csv("failed_emails.csv", index=False)
    else:
        print("\n✅ All emails sent successfully!")

if __name__ == "__main__":
    run_app()
