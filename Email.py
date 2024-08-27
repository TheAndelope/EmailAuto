import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
import os
import openai

load_dotenv()

openai.api_key = os.getenv("OPENAI_API_KEY")

df = pd.read_csv('sponsors.csv')

your_email = os.getenv("EMAIL")
your_password = os.getenv("PASS")

smtp_server = "smtp.gmail.com"
smtp_port = 587

server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(your_email, your_password)



def generate_text(company, model="gpt-4o-mini", max_tokens=350):
    messages = [{"role": "system", "content": "You are a member of a tech startup called the Neo Developer League and are explaining to a company why your company's values align with theirs and be specific. The general idea of Neo Dev is : Neo Developer League is a student-led organization that hosts competitive events created to inspire high school students to pursue engineering and build connections in a fun and competitive way. Also limit answer to a paragraph"}]
    messages.append({"role": "user", "content": f'Tell me why your company algins with the values at {company}'})

    response = openai.chat.completions.create(
        model=model,
        messages=messages,
        max_tokens=max_tokens,
        temperature=0.7,
        n=1,
        stop=None,
        stream=False,
        frequency_penalty = 2.0
    )
    return response.choices[0].message.content

for index, row in df.iterrows():
    status = row['Status']
    company_name = row['Company Name']
    first_name = row['First Name']
    recipient_email = row['Email']
    subject = f"Why {company_name} Should Sponsor Us"
    body = f"""
    Dear {first_name},
    
    I hope this message finds you well. We at the Neo Developer League are reaching out to {company_name} because we believe your company aligns with our values and vision for the future. 
    
    {generate_text(company={company_name})}
    
    Thank you for considering our request, and we look forward to the possibility of working together!
    
    Best regards,
    Andy Duong
    Finance Lead
    theandelope16@gmail.com
    """
    
    msg = MIMEMultipart()
    msg['From'] = your_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    server.sendmail(your_email, recipient_email, msg.as_string())
    print(f"Email sent to {first_name} at {company_name}")


server.quit()
