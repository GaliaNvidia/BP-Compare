# Email Setup Guide

## How to Send Results via Email with Excel Attachment

The app now supports sending comparison results directly via email with the Excel file attached.

## Quick Setup

### Step 1: Configure Email Settings in Sidebar

1. Open the app sidebar (click `>` if collapsed)
2. Enter the recipient email (default: galiaf@nvidia.com)
3. Click **"⚙️ SMTP Settings"** to expand
4. Enter your email configuration:

### Step 2: NVIDIA Office 365 Settings

For NVIDIA email accounts:

| Setting | Value |
|---------|-------|
| **SMTP Server** | `smtp.office365.com` |
| **SMTP Port** | `587` |
| **Your Email** | Your full NVIDIA email (e.g., yourname@nvidia.com) |
| **Password** | Your email password or app-specific password |

### Step 3: Send Email

1. Upload and compare your Excel files
2. Click **"📧 Send Email with Results"** button
3. The email will be sent with:
   - Summary of changes in the body
   - Excel file attached
   - Professional formatting

## Security Notes

### Using App Passwords (Recommended)

For better security, use an **app-specific password** instead of your main password:

#### For Microsoft 365:
1. Go to [Microsoft Account Security](https://account.microsoft.com/security)
2. Navigate to "Advanced security options"
3. Under "App passwords", click "Create a new app password"
4. Copy the generated password
5. Use this password in the app

### Alternative: Use Your Regular Password

If app passwords aren't enabled, you can use your regular NVIDIA email password. However:
- ⚠️ The password is only stored in memory during the session
- 🔒 It's never saved to disk
- 🔄 You'll need to re-enter it each time you restart the app

## Troubleshooting

### "Authentication failed" Error

**Causes:**
- Incorrect email or password
- Multi-factor authentication (MFA) blocking
- Account security settings

**Solutions:**
1. Double-check your email address is complete
2. Use an app-specific password instead
3. Verify MFA isn't blocking SMTP access
4. Contact IT if issues persist

### "Connection timeout" Error

**Causes:**
- Firewall blocking port 587
- VPN interference
- SMTP server unavailable

**Solutions:**
1. Check your network connection
2. Try disconnecting from VPN
3. Verify SMTP server: `smtp.office365.com`
4. Ensure port 587 is open

### "Recipient rejected" Error

**Causes:**
- Invalid recipient email address
- Email blocked by recipient's settings

**Solutions:**
1. Verify recipient email is correct
2. Ensure recipient accepts emails from your address

## Email Content

The email will contain:

**Subject:** `PO Comparison Results - YYYY-MM-DD`

**Body:**
```
Hi,

Please find attached the PO Line Comparison results:

SUMMARY:
- Total changes: X
- Pushed lines: Y
- Split lines: Z
- Alerts (>7 days): W

Generated: YYYY-MM-DD HH:MM:SS

Best regards,
PO Line Comparison Tool
```

**Attachment:** `po_comparison_YYYYMMDD_HHMMSS.xlsx`

## Alternative: Manual Email

If automatic email doesn't work:

1. Click **"📥 Download Results as Excel"**
2. Open your email client manually
3. Compose email to galiaf@nvidia.com
4. Attach the downloaded Excel file
5. Send

## Privacy & Security

- ✅ Passwords are only stored in session memory
- ✅ No data is saved to disk
- ✅ Connection uses TLS encryption (STARTTLS)
- ✅ No third-party services involved
- ⚠️ Consider using app passwords instead of main password

## Support

For issues or questions:
- Check the main README.md
- Verify SMTP settings with IT
- Test email settings in your regular email client first

