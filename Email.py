'''
Auto Email Program
By: Andy Duong
//Description\\

Actions
y -> sends email
r -> regenerates email
e -> edits email
s -> skips entry
h -> sets status to handwrite (indicates the email must be handwritten)


Note** When you edit a gui window will apear in a separate window so you might have to select it in taskbar

'''
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
import os
import openai
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import tkinter as tk
from tkinter import scrolledtext
from datetime import datetime
from email.utils import formataddr

# Load environment variables
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")
your_email = os.getenv("EMAIL")
sender_name = 'NeoDev League'
your_password = os.getenv("PASS")

# Set up Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheet = client.open("Email Test").sheet1
data = sheet.get_all_records()
df = pd.DataFrame(data)

# Set up SMTP server
smtp_server = "smtp.gmail.com"
smtp_port = 587
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(your_email, your_password)

name = ""
while True:
    name = input("What Is Your Full Name (Your name WILL be included on the current email template)? ")
    if name.strip() != "":
        choice = input(f'Are you sure {name} is your full real name (y/n) ').strip().lower()
        if choice == 'y':
            break

def generate_text(company, model="gpt-4o-mini", max_tokens=300):
    messages = [{"role": "system", "content": "You are a member of a tech startup called the Neo Developer League and are explaining to a company why your company's values align with theirs and be specific. The general idea of the Neo Developer League is: the Neo Developer League is a student-led organization that hosts competitive events created to inspire high school students to pursue engineering and build connections in a fun and competitive way. Also limit answer to ONE paragraph and say it like you just finished explaining the company. Also speak as a team (use 'we)"}]
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

def generate_email(company_name, first_name):
    body = f"""
<html>
<head></head>
<body>
<p>{first_name},</p>
<p>Hello! I'm {name}, a founder at the NeoDev League, and we're reaching out to collaborate! We’re excited to connect with you as we prepare to launch our first competitive engineering event for high school students this October. We’re expecting 150-200 participants for this one-day event, where teams of students will collaborate on projects to present to a panel of judges. Our organization is focused on sparking collaboration and innovation in high schoolers. Events will allow teams of 8-10 students from various schools in the region to collaborate over an 8-hour work period to create a project and present it to a panel of judges.</p>
<p>{generate_text(company=company_name)}</p>
<p>We’d love to discuss this further and explore how we can collaborate. Please feel free to reach out with any questions.</p>
<p>Thank you for your time and consideration.</p>
<table cellpadding="0" cellspacing="0" border="0" globalstyles="[object Object]" class="table__StyledTable-sc-1avdl6r-0 lmHSv" style="vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial;">
    <tbody>
        <tr>
            <td>
                <table cellpadding="0" cellspacing="0" border="0" globalstyles="[object Object]" class="table__StyledTable-sc-1avdl6r-0 lmHSv" style="vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial;">
                    <tbody>
                        <tr>
                            <td style="vertical-align: top;">
                                <table cellpadding="0" cellspacing="0" border="0" globalstyles="[object Object]" class="table__StyledTable-sc-1avdl6r-0 lmHSv" style="vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial;">
                                    <tbody>
                                        <tr>
                                            <td class="template1__ImageContainer-sc-nmby7a-0 byHTIl" style="text-align: center;">
                                                <img src="https://cdn.discordapp.com/attachments/1240800319043272828/1282369090098565201/upscaled_image-removebg.png?ex=66df1ac8&amp;is=66ddc948&amp;hm=8eccffb4ce1c5a93920b89dc0c621ca87ca2d30b5def2af590f8d92d280f1af7&amp;" role="presentation" width="130" class="image__StyledImage-sc-hupvqm-0 fOIYAq" style="display: block; max-width: 130px;">
                                            </td>
                                        </tr>
                                        <tr>
                                            <td height="30"></td>
                                        </tr>
                                        <tr>
                                            <td style="text-align: center;">
                                                <table cellpadding="0" cellspacing="0" border="0" globalstyles="[object Object]" class="table__StyledTable-sc-1avdl6r-0 lmHSv" style="display: inline-block; vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial;">
                                                    <tbody>
                                                        <tr style="text-align: center;">
                                                            <td>
                                                                <a href="https://x.com/NeoDevLeague" color="#065f46" class="social-links__LinkAnchor-sc-py8uhj-2 iWFMIm" style="display: inline-block; padding: 0px; background-color: rgb(6, 95, 70);">
                                                                    <img src="https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/twitter-icon-2x.png" alt="twitter" color="#065f46" width="24" class="social-links__LinkImage-sc-py8uhj-1 cgBQhD" style="background-color: rgb(6, 95, 70); max-width: 135px; display: block;">
                                                                </a>
                                                            </td>
                                                            <td width="5">
                                                                <div></div>
                                                            </td>
                                                            <td>
                                                                <a href="https://www.linkedin.com/company/neo-developer-league" color="#065f46" class="social-links__LinkAnchor-sc-py8uhj-2 iWFMIm" style="display: inline-block; padding: 0px; background-color: rgb(6, 95, 70);">
                                                                    <img src="https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/linkedin-icon-2x.png" alt="linkedin" color="#065f46" width="24" class="social-links__LinkImage-sc-py8uhj-1 cgBQhD" style="background-color: rgb(6, 95, 70); max-width: 135px; display: block;">
                                                                </a>
                                                            </td>
                                                            <td width="5">
                                                                <div></div>
                                                            </td>
                                                            <td>
                                                                <a href="https://www.instagram.com/neodevleague" color="#065f46" class="social-links__LinkAnchor-sc-py8uhj-2 iWFMIm" style="display: inline-block; padding: 0px; background-color: rgb(6, 95, 70);">
                                                                    <img src="https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/instagram-icon-2x.png" alt="instagram" color="#065f46" width="24" class="social-links__LinkImage-sc-py8uhj-1 cgBQhD" style="background-color: rgb(6, 95, 70); max-width: 135px; display: block;">
                                                                </a>
                                                            </td>
                                                            <td width="5">
                                                                <div></div>
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </td>
                                        </tr>
                                    </tbody>
                                </table>
                            </td>
                            <td width="46">
                                <div></div>
                            </td>
                            <td style="padding: 0px; vertical-align: middle;">
                                <h2 color="#000000" class="name__NameContainer-sc-1m457h3-0 csBPEs" style="margin: 0px; font-size: 18px; color: rgb(0, 0, 0); font-weight: 600;">
                                    <span>The NeoDev</span><span>&nbsp;</span><span>Team</span>
                                </h2>
                                <p color="#000000" font-size="medium" class="company-details__CompanyContainer-sc-j5pyy8-0 cSOAsl" style="margin: 0px; font-weight: 500; color: rgb(0, 0, 0); font-size: 14px; line-height: 22px;">
                                    <span>Neo Developer League</span>
                                </p>
                                <table cellpadding="0" cellspacing="0" border="0" globalstyles="[object Object]" class="table__StyledTable-sc-1avdl6r-0 lmHSv" style="width: 100%; vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial;">
                                    <tbody>
                                        <tr>
                                            <td height="30"></td>
                                        </tr>
                                        <tr>
                                            <td color="#065f46" direction="horizontal" width="auto" height="1" class="color-divider__Divider-sc-1h38qjv-0 bofWVx" style="width: 100%; border-bottom: 1px solid rgb(6, 95, 70); border-left: none; display: block;"></td>
                                        </tr>
                                        <tr>
                                            <td height="30"></td>
                                        </tr>
                                    </tbody>
                                </table>
                                <table cellpadding="0" cellspacing="0" border="0" globalstyles="[object Object]" class="table__StyledTable-sc-1avdl6r-0 lmHSv" style="vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial;">
                                    <tbody>
                                        <tr height="25" style="vertical-align: middle;">
                                            <td width="30" style="vertical-align: middle;">
                                                <table cellpadding="0" cellspacing="0" border="0" globalstyles="[object Object]" class="table__StyledTable-sc-1avdl6r-0 lmHSv" style="vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial;">
                                                    <tbody>
                                                        <tr>
                                                            <td style="vertical-align: bottom;">
                                                                <span color="#065f46" width="11" class="contact-info__IconWrapper-sc-mmkjr6-1 ldYaqt" style="display: inline-block; background-color: rgb(6, 95, 70);">
                                                                    <img src="https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/email-icon-2x.png" color="#065f46" alt="emailAddress" width="13" class="contact-info__ContactLabelIcon-sc-mmkjr6-0 gxFfYp" style="display: block; background-color: rgb(6, 95, 70);">
                                                                </span>
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </td>
                                            <td style="padding: 0px;">
                                                <a href="mailto:info@neoleague.dev" color="#000000" class="contact-info__ExternalLink-sc-mmkjr6-2 jOTYAn" style="text-decoration: none; color: rgb(0, 0, 0); font-size: 12px;">
                                                    <span>info@neoleague.dev</span>
                                                </a>
                                            </td>
                                        </tr>
                                        <tr height="25" style="vertical-align: middle;">
                                            <td width="30" style="vertical-align: middle;">
                                                <table cellpadding="0" cellspacing="0" border="0" globalstyles="[object Object]" class="table__StyledTable-sc-1avdl6r-0 lmHSv" style="vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial;">
                                                    <tbody>
                                                        <tr>
                                                            <td style="vertical-align: bottom;">
                                                                <span color="#065f46" width="11" class="contact-info__IconWrapper-sc-mmkjr6-1 ldYaqt" style="display: inline-block; background-color: rgb(6, 95, 70);">
                                                                    <img src="https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/link-icon-2x.png" color="#065f46" alt="website" width="13" class="contact-info__ContactLabelIcon-sc-mmkjr6-0 gxFfYp" style="display: block; background-color: rgb(6, 95, 70);">
                                                                </span>
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </td>
                                            <td style="padding: 0px;">
                                                <a href="https://neoleague.dev" color="#000000" class="contact-info__ExternalLink-sc-mmkjr6-2 jOTYAn" style="text-decoration: none; color: rgb(0, 0, 0); font-size: 12px;">
                                                    <span>neoleague.dev</span>
                                                </a>
                                            </td>
                                        </tr>
                                    </tbody>
                                </table>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </td>
        </tr>
    </tbody>
</table>
</body>
</html>
    """
    return body

def edit_email_body(initial_body):
    def save_and_close():
        global email_body
        email_body = text_area.get("1.0", tk.END).strip()
        root.destroy()

    root = tk.Tk()
    root.title("Edit Email Body")

    text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=20)
    text_area.insert(tk.INSERT, initial_body)
    text_area.pack(padx=10, pady=10)

    save_button = tk.Button(root, text="Save and Close", command=save_and_close)
    save_button.pack(pady=10)

    root.mainloop()
    return email_body

def send_email(recipient_email, subject, body):
    msg = MIMEMultipart('alternative')
    msg['From'] =formataddr((sender_name, your_email))
    msg['To'] = recipient_email
    msg['Cc'] = "admin@neoleague.dev"
    msg['Subject'] = subject
    
    part = MIMEText(body, 'html')
    msg.attach(part)

    file_path = 'NeoDev_Sponsorship_Package.pdf'
    try:
        with open(file_path, 'rb') as file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file_path)}')
            msg.attach(part)
    except FileNotFoundError:
        for i in range(0, 100):
            print(f"DONT WORRY NO EMAIL WAS SENT! File {file_path} not found. You may have named the Sponsorship package incorrectly, check name with line 154 (case sensitive) or ask Andy. Skip entry for now if needed")
            print('\n')
        quit()
    
    server.sendmail(your_email, recipient_email, msg.as_string())

for index, row in df.iterrows():
    status = row['Status']
    if status != "Selected":
        continue

    df.at[index, 'Status'] = 'In Process'
    cell = sheet.find(row['Company Name'])
    sheet.update_cell(cell.row, df.columns.get_loc('Status') + 1, 'In Process')

    company_name = row['Company Name']
    first_name = row['First Name']
    if not (isinstance(first_name, str)) or first_name.strip() == '':
        first_name = "To Whom It May Concern"
    else:
        first_name = "Dear " + first_name

    recipient_email = row['Email']
    subject = f"Why {company_name} Should Sponsor Us"
    body = generate_email(company_name=company_name, first_name=first_name)
    
    '''
    print(f"\nSubject: {subject}")
    print(f"To: {recipient_email}")
    print(f"Body:\n{body}")
    '''
    #auto email
    
    try:
        send_email(recipient_email, subject, body)
        df.at[index, 'Status'] = 'Review'
        print(f"Email sent to {first_name} at {company_name}")
        cell = sheet.find(row['Company Name'])
        sheet.update_cell(cell.row, df.columns.get_loc('Status') + 1, 'Review')
        sheet.update_cell(cell.row, df.columns.get_loc('Date Sent') + 1, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        sheet.update_cell(cell.row, df.columns.get_loc('Person') + 1, name)
    except Exception as e:
        print(f"Failed to send email: {e}")
    

    #manual email
    '''
    while True:
        send_email_choice = input("Do you want to send this email? (yes (y) | edit (e) | regenerate (r) | skip (s) | handwrite (h)): ").strip().lower()
        
        if send_email_choice in ('y', 'yes'):
            send_email_choice = input("Are you sure? (y/n): ").strip().lower()
            if send_email_choice in ('y', 'yes'):
                try:
                    send_email(recipient_email, subject, body)
                    df.at[index, 'Status'] = 'Review'
                    print(f"Email sent to {first_name} at {company_name}")
                    cell = sheet.find(row['Company Name'])
                    sheet.update_cell(cell.row, df.columns.get_loc('Status') + 1, 'Review')
                    sheet.update_cell(cell.row, df.columns.get_loc('Date Sent') + 1, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
                    sheet.update_cell(cell.row, df.columns.get_loc('Person') + 1, name)
                except Exception as e:
                    print(f"Failed to send email: {e}")
                break
            else:
                pass

        elif send_email_choice in ('e', 'edit'):
            body = edit_email_body(body)
            print(f"\nSubject: {subject}")
            print(f"To: {recipient_email}")
            print(f"Body:\n{body}")

        elif send_email_choice in ('r', 'regenerate'):
            body = generate_email(company_name=company_name, first_name=first_name)
            print(f"\nSubject: {subject}")
            print(f"To: {recipient_email}")
            print(f"Body:\n{body}")

        elif send_email_choice in ('s', 'skip'):
            if status == 'In Process':
                df.at[index, 'Status'] = 'Selected'
                cell = sheet.find(row['Company Name'])
                sheet.update_cell(cell.row, df.columns.get_loc('Status') + 1, 'Selected')
            break

        elif send_email_choice in ('h', 'handwrite'):
            df.at[index, 'Status'] = 'Handwrite'
            cell = sheet.find(row['Company Name'])
            sheet.update_cell(cell.row, df.columns.get_loc('Status') + 1, 'Handwrite')
            break

        else:
            print("Invalid input. Please enter 'y' to send or 'n' to create another email.")
    '''

df.to_csv('sponsors.csv', index=False)
server.quit()
