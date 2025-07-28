import smtplib
import os
import pandas as pd
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

load_dotenv()

# Email Configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
RESUME_FILE = "Awadhesh_Sharma_Resume.pdf"

# Ensure resume exists
if not os.path.exists(RESUME_FILE):
    print(f"❌ Resume file '{RESUME_FILE}' not found. Please place it in the script folder.")
    exit()

# List to track failed emails
failed_log = []

def send_confirmation_email(email, company, name):
    """Sends confirmation email with HTML body and resume attachment."""
    subject = f"Regarding Job Opportunity at {company}"

    # HTML body
    body = f"""\
Hi {name},<br><br>

I hope you're doing well.<br><br>

I’m reaching out to express my interest in a <strong>Frontend Developer</strong> position at <strong>{company}</strong>.<br><br>

With over 3 years of experience in Angular, I’ve built scalable web applications using:
<ul>
  <li><strong>Angular (v13+)</strong> with standalone components</li>
  <li><strong>Reactive Forms</strong> for dynamic form control</li>
  <li><strong>NgRx</strong> for advanced state management</li>
  <li>DevExtreme, Highcharts, Power BI for dynamic dashboards</li>
</ul>

I believe my technical skills and problem-solving mindset would be a great fit for your team.<br><br>

Please find my resume attached. Looking forward to the opportunity to connect.<br><br>

<p style="margin-bottom: 6px;">Regards,</p>

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
        # Create multipart email
        msg = MIMEMultipart()
        msg["From"] = EMAIL_SENDER
        msg["To"] = email
        msg["Subject"] = subject

        msg.attach(MIMEText(body, "html"))

        # Attach resume
        with open(RESUME_FILE, "rb") as f:
            part = MIMEApplication(f.read(), Name=RESUME_FILE)
        part['Content-Disposition'] = f'attachment; filename="{RESUME_FILE}"'
        msg.attach(part)

        # Send email
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
        data = pd.read_csv("Hr-detail.csv")
    except FileNotFoundError:
        print("❌ 'Hr-detail.csv' file not found.")
        return

    for index, row in data.iterrows():
        name = str(row.get('Name', '')).strip()
        email = str(row.get('Email', '')).strip()
        company = str(row.get('Company', '')).strip()

        if not email or not name or not company:
            print(f"⚠️ Skipping incomplete row {index + 2} in CSV")
            continue

        send_confirmation_email(email, company, name)

    if failed_log:
        print(f"\n❗ {len(failed_log)} email(s) failed. Logging to failed_emails.csv")
        pd.DataFrame(failed_log).to_csv("failed_emails.csv", index=False)
    else:
        print("\n✅ All emails sent successfully!")

if __name__ == "__main__":
    run_app()
