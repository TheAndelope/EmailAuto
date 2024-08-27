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
    messages = [{"role": "system", "content": "You are a member of a tech startup called the Neo Developer League and are explaining to a company why your company's values align with theirs and be specific. The general idea of Neo Dev is : Neo Developer League is a student-led organization that hosts competitive events created to inspire high school students to pursue engineering and build connections in a fun and competitive way. Also limit answer to ONE paragraph"}]
    messages.append({"role": "user", "content": f'Tell me why your company aligns with the values at {company}'})

    response = openai.chat.completions.create(
        model=model,
        messages=messages,
        max_tokens=max_tokens,
        temperature=0.7,
        n=1,
        stop=None,
        stream=False,
        frequency_penalty=2.0
    )
    return response.choices[0].message.content

for index, row in df.iterrows():
    status = row['Status']
    company_name = row['Company Name']
    first_name = row['First Name']

    if first_name.strip() == '':
        first_name = "To Whom It May Concern"
    else:
        first_name = "Dear, " + first_name

    recipient_email = row['Email']
    subject = f"Why {company_name} Should Sponsor Us"
    body = f"""
{first_name},

Hello! I'm Andy Duong, a founder at the NeoDev League, and we're reaching out to collaborate! We’re excited to connect with you as we prepare to launch our first competitive engineering event for high school students this October. We’re expecting 150-200 participants for this one-day event, where teams of students will collaborate on projects to present to a panel of judges. Our organization is focused on sparking collaboration and innovation in high schoolers. Events will allow teams of 8-10 students from various schools in the region to collaborate over an 8 hour work period to create a project and present it to a panel of judges.  

{generate_text(company=company_name)}

We’d love to discuss this further and explore how we can collaborate. Please feel free to reach out with any questions.

Thank you for your time and consideration.

Best regards,
Andy Duong
Finance Lead
neodevleague@gmail.com
    """

    while True:
        print(f"\nSubject: {subject}")
        print(f"To: {recipient_email}")
        print(f"Body:\n{body}")

        send_email = input("Do you want to send this email? (y/n): ").strip().lower()
        
        if send_email == 'y':
            msg = MIMEMultipart()
            msg['From'] = your_email
            msg['To'] = recipient_email
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))
            
            server.sendmail(your_email, recipient_email, msg.as_string())
            print(f"Email sent to {first_name} at {company_name}")
            break
        elif send_email == 'n':
            print("Here's Another")
            break
        else:
            print("Invalid input. Please enter 'y' to send or 'n' to modify.")
            
server.quit()
