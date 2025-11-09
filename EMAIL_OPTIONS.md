# Email Options Guide

## 🎉 No Password Required Options!

You now have **3 different ways** to send emails with the Excel attachment. Two of them **don't require any password configuration**!

---

## Option 1: Outlook (Mac) - RECOMMENDED ✨

### ✅ Advantages:
- **No password needed!**
- Direct integration with Outlook
- Creates a draft email automatically
- Excel file already attached
- Review before sending

### How it works:
1. Select **"Outlook (Mac)"** in the sidebar
2. Upload and compare your files
3. Click **"Create Outlook Draft"**
4. Outlook opens with a draft email containing:
   - Pre-filled recipient (galiaf@nvidia.com)
   - Subject line
   - Email body with summary
   - **Excel file already attached!**
5. Review and click Send in Outlook

### Requirements:
- Microsoft Outlook must be installed on your Mac
- Outlook must be configured with your NVIDIA account

---

## Option 2: Save as .eml File - MOST COMPATIBLE 📧

### ✅ Advantages:
- **No password needed!**
- Works with ANY email client (Outlook, Mail, Gmail, Thunderbird)
- Works on Mac, Windows, and Linux
- Excel file embedded in the .eml file
- Can be saved and opened later

### How it works:
1. Select **"Save as .eml file"** in the sidebar
2. Upload and compare your files
3. Click **"Download .eml File"**
4. A .eml file is downloaded to your computer
5. **Double-click the .eml file** to open it in your default email client
6. The email opens with:
   - Pre-filled recipient (galiaf@nvidia.com)
   - Subject line
   - Email body with summary
   - **Excel file already attached!**
7. Review and click Send

### Compatible with:
- ✅ Microsoft Outlook (Mac & Windows)
- ✅ Apple Mail (Mac)
- ✅ Gmail (via desktop client)
- ✅ Thunderbird
- ✅ Windows Mail
- ✅ Any RFC822-compliant email client

---

## Option 3: SMTP - Direct Send ⚙️

### ✅ Advantages:
- Sends email immediately
- No manual steps after clicking send
- Email sent directly from the app

### ⚠️ Disadvantages:
- **Requires email password**
- May need app-specific password
- More complex setup

### How it works:
1. Select **"SMTP"** in the sidebar
2. Expand **"SMTP Settings"** and enter:
   - SMTP Server: `smtp.office365.com`
   - Port: `587`
   - Your Email: Your NVIDIA email address
   - Password: Your email password or app password
3. Upload and compare your files
4. Click **"Send Email via SMTP"**
5. Email is sent immediately!

### Requirements:
- Your NVIDIA email credentials
- May need app-specific password for MFA accounts
- Network access to SMTP server

---

## Comparison Table

| Feature | Outlook (Mac) | .eml File | SMTP |
|---------|--------------|-----------|------|
| **Password Required** | ❌ No | ❌ No | ✅ Yes |
| **Platform** | Mac only | All platforms | All platforms |
| **Excel Attached** | ✅ Auto | ✅ Auto | ✅ Auto |
| **Review Before Send** | ✅ Yes | ✅ Yes | ❌ No |
| **Setup Required** | Outlook app | None | Email credentials |
| **Immediate Send** | ❌ No (draft) | ❌ No (manual) | ✅ Yes |
| **Works Offline** | ✅ Yes | ✅ Yes | ❌ No |

---

## Which Option Should I Use?

### Use **Outlook (Mac)** if:
- ✅ You're on a Mac
- ✅ You have Outlook installed
- ✅ You want the easiest option
- ✅ You want to review before sending

### Use **.eml File** if:
- ✅ You don't have Outlook on Mac
- ✅ You're on Windows or Linux
- ✅ You use any email client
- ✅ You want maximum compatibility
- ✅ You don't want to enter passwords

### Use **SMTP** if:
- ✅ You want immediate sending
- ✅ You don't want any manual steps
- ✅ You're comfortable with email passwords
- ✅ You want fully automated email

---

## Recommended: .eml File Method 🌟

For most users, we recommend the **.eml file** method because:

1. **No password required** 🔐
2. **Works everywhere** 💻
3. **Works with any email client** 📧
4. **Simple and safe** ✅
5. **You can review before sending** 👀

---

## Step-by-Step: Using .eml File (Recommended)

1. **In the app sidebar**, select **"Save as .eml file - No password needed! 📧"**

2. **Upload your Excel files** and run the comparison

3. **Click "Download .eml File (with attachment)"**
   - A file like `po_email_20241109_143022.eml` will be downloaded

4. **Find the downloaded .eml file** (usually in your Downloads folder)

5. **Double-click the .eml file**
   - Your default email client will open
   - You'll see a new email window with:
     - To: galiaf@nvidia.com
     - Subject: PO Comparison Results - [date]
     - Body: Summary of changes
     - Attachment: Excel file with full results

6. **Review the email** and click **Send**!

That's it! No passwords, no configuration, works everywhere! 🎉

---

## Troubleshooting

### Outlook (Mac) doesn't work
- Verify Outlook is installed: Open Applications → Microsoft Outlook
- Ensure Outlook is configured with your NVIDIA account
- Try the .eml file method instead

### .eml file opens but no attachment
- This is rare but can happen with some email clients
- Solution: Use the "Download Results as Excel" button separately
- Manually attach the Excel file to the email

### SMTP authentication fails
- Try using an app-specific password
- Verify your email and password are correct
- Check if MFA is blocking SMTP access
- Consider using .eml file method instead

---

## Security Notes

### Password Safety
- The app **never saves passwords** to disk
- Passwords are only kept in memory during your session
- When you close the app, passwords are forgotten
- For SMTP, consider using app-specific passwords

### Email Content
- All email content is generated locally
- No data is sent to external services
- The Excel file is embedded directly in the .eml file
- Everything stays under your control

---

Need help? Check the main README.md or EMAIL_SETUP.md for more details!

