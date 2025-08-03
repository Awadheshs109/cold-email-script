import os
import smtplib
import time
import random
import pandas as pd
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

load_dotenv()

# Email config
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
RESUME_LINK = "https://drive.google.com/file/d/17fpc_EwLpJCKKFGPfzZiHuOamPHaILri/view?usp=sharing"

# Error log path
LOG_FILE = "error_log.txt"

# Function to send email
def send_confirmation_email(email, company, name):
    subject = f"Regarding Job Opportunity at {company}"
    body = f"""\
Hi {name},<br><br>

I hope you're doing well.
I’m reaching out to express my interest in the <strong>Frontend Angular Developer</strong> position at <strong>{company}</strong>.<br>
With over 3.5 years of experience in Angular, I have built scalable web applications using:
<ul>
  <li><strong>Angular (v13)</strong> with standalone components</li>
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
        return True

    except Exception as e:
        error_msg = f"❌ Failed to send email to {name} ({email}) - {e}\n"
        print(error_msg)
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(error_msg)
        return False

def run_app():
    file_path = "HR-Database.xlsx"
    if not os.path.exists(file_path):
        print(f"❌ '{file_path}' not found.")
        return

    print("📄 Reading contacts from Excel...")
    df_all = pd.read_excel(file_path)

    # Add 'Email Sent' column if not exists
    if 'Email Sent' not in df_all.columns:
        df_all['Email Sent'] = ""

    # Work only on rows 101 to 400 (index 100 to 399)
    df_slice = df_all.iloc[0:250].copy()

    # Drop rows in the slice that don't have required fields
    df_slice = df_slice.dropna(subset=["Email", "Name", "Company"])

    print(f"📧 Sending emails to rows 0–250 ({len(df_slice)} contacts)...")

    for i in df_slice.index:
        row = df_all.loc[i]
        name = str(row.get("Name", "")).strip()
        email = str(row.get("Email", "")).strip()
        company = str(row.get("Company", "")).strip()
        already_sent = str(row.get("Email Sent", "")).strip().lower()

        if already_sent == "yes":
            print(f"⏭️ Skipping {name} ({email}) — already sent.")
            continue

        if not email or not name or not company or "@" not in email:
            print(f"⚠️ Skipping invalid row {i + 1}")
            df_all.at[i, "Email Sent"] = "No"
            continue

        print(f"📨 Sending for row {i + 1}...")
        sent = send_confirmation_email(email, company, name)
        df_all.at[i, "Email Sent"] = "Yes" if sent else "No"

        delay = random.randint(5, 10)
        print(f"⏳ Waiting {delay}s before next...")
        time.sleep(delay)

    # Save updated full DataFrame back
    df_all.to_excel(file_path, index=False)
    print("\n✅ Script finished and Excel updated.")

if __name__ == "__main__":
    run_app()
