import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

df = pd.read_csv('sponsors.csv')

your_email = "your_email@example.com"
your_password = "your_password"

smtp_server = "smtp.gmail.com"
smtp_port = 587

server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()  # Secure the connection
server.login(your_email, your_password)

for index, row in df.iterrows():
    status = row['status']
    company_name = row['company name']
    first_name = row['first name']
    recipient_email = row['email']
    subject = f"Why {company_name} Should Sponsor Us"
    body = f"""
    Dear {first_name},
    
    I hope this message finds you well. We at [Your Organization] are reaching out to {company_name} because we believe your company aligns with our values and vision for the future. 
    We would love to discuss potential sponsorship opportunities and how your support could make a significant impact on our upcoming projects.
    
    Thank you for considering our request, and we look forward to the possibility of working together!
    
    Best regards,
    [Your Name]
    [Your Position]
    [Your Contact Information]
    """
    
    msg = MIMEMultipart()
    msg['From'] = your_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    server.sendmail(your_email, recipient_email, msg.as_string())
    print(f"Email sent to {first_name} at {company_name}")


server.quit()
