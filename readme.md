# 📧 Auto Email Program

Welcome to the **Auto Email Program**! This application automatically sends personalized emails using data from a Google Spreadsheet. With a few simple actions, you can generate, edit, and send emails to your contacts. This program is designed to streamline your communication, especially for sponsorship outreach. 🚀

---

## 🔧 Features
- **Automated email generation** using GPT-based templates.
- **Personalized email content** based on company names.
- **Edit and regenerate** emails on the go.
- **Attach sponsorship packages** to emails.
- **Track status** updates directly in Google Sheets.

---

## 🚀 How It Works

### Actions:
| Key | Action |
| --- | ------ |
| `y` | Sends email 📧 |
| `r` | Regenerates email 🔄 |
| `e` | Edits email 📝 |
| `s` | Skips entry ⏩ |
| `h` | Marks email to be handwritten ✍️ |

When editing, a separate GUI window will pop up, which you may need to select from the taskbar.

---

## ⚙️ Setup Instructions

1. **Clone the repository**:
   ```bash
   git clone https://github.com/your-repo/auto-email-program.git
   cd auto-email-program
   ```

2. **Install dependencies**:
   Make sure you have Python 3 installed. Then install required packages:
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up environment variables**:
   Create a `.env` file in the root directory:
   ```env
   EMAIL=your_email@example.com
   PASS=your_email_password
   OPENAI_API_KEY=your_openai_api_key
   ```

4. **Configure Google Sheets**:
   - Set up Google Sheets API and download the `credentials.json` file.
   - Make sure to have a spreadsheet named **"Email Test"** with columns: `Company Name`, `First Name`, `Email`, `Status`, `Date Sent`, and `Person`.

---

## ✨ Usage

1. **Start the program**:
   ```bash
   python auto_email.py
   ```

2. **Enter your full name**:
   Your name will be included in the email template. Confirm your input.

3. **Email generation and sending**:
   - The program will generate and display the email.
   - You can choose to send the email, edit, regenerate, skip, or mark for handwritten follow-up.

4. **Email status updates**:
   The program will update the `Status`, `Date Sent`, and `Person` fields in the Google Spreadsheet after each email is processed.

---

## 📝 Customization

- **Email Templates**:
  The default email template is customizable in the `generate_email` function. Adjust the body to suit your needs.
  
- **Attachments**:
  Ensure the sponsorship package or any other file you want to send is named correctly (`NeoDev_Sponsorship_Package.pdf`).

---

## 🛠️ Troubleshooting

- **File Not Found**: If the attachment is missing, the program will notify you and ask you to skip the entry.
- **Email Errors**: If an email fails to send, the program will log the error for you to review.

---

Enjoy your automated emailing experience! 🎉 If you encounter any issues, feel free to reach out!
