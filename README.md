# Gmail/Sheets Mail Merge

Enhanced mail merge script for Google Apps Script that sends personalized emails using Gmail drafts as templates and Google Sheets for recipient data.

## Features

- Use Gmail draft messages as email templates
- **Three operation modes**: Draft creation, MailApp sending, GmailApp sending
- **Smart emoji conversion** - automatically handles emojis in all modes  
- **Draft mode** - create drafts for manual review before sending
- **MailApp mode** - best emoji support with immediate sending
- **GmailApp mode** - full Gmail features with immediate sending
- Multiple recipients per row (comma-separated format)
- Template variable substitution using `{{variable}}` syntax
- File attachments and inline images support (draft/GmailApp modes)
- Send-as aliases and Gmail advanced features (draft/GmailApp modes)
- Email tracking with sent/draft status column
- Error handling and logging

## Setup

1. Open Google Sheets and create a new spreadsheet
2. Set up your data with column headers including:
   - `Recipient`: Email addresses (can be comma-separated for multiple recipients)
   - `Email Sent`: Status column (leave blank initially)
   - Any other columns for template variables (e.g., `Name`, `Company`, etc.)

3. Go to Extensions → Apps Script
4. Delete the default code and paste the contents of `mail-merge.gs`
5. Save the project

## Usage

### 1. Create Email Template
1. In Gmail, create a new draft message
2. Write your email using `{{ColumnName}}` syntax for variables
3. Example: "Hi {{Name}}, thanks for your interest in {{Company}}!"
4. Save as draft (don't send)

### 2. Set Up Spreadsheet Data
Example spreadsheet format:
```
Recipient                           | Name    | Company | Email Sent
user1@example.com                   | John    | Acme    |
user2@example.com,user3@example.com | Jane    | Corp    |
```

### 3. Choose Your Mail Merge Mode

The script offers three different modes to fit your needs:

#### Option A: Draft Emails (Recommended for Review)
1. Click "Mail Merge" → "Draft Emails"
2. Enter the subject line of your Gmail draft template  
3. Script creates personalized drafts in your Gmail drafts folder
4. Review each draft manually before sending
5. Spreadsheet shows "Draft created [timestamp]" in the status column

#### Option B: Send with MailApp (Best Emoji Support)
1. Click "Mail Merge" → "Send Emails (MailApp - best emoji support)"
2. Enter the subject line of your Gmail draft
3. Script sends emails immediately using MailApp
4. Spreadsheet shows timestamp in the "Email Sent" column

#### Option C: Send with GmailApp (Full Gmail Features)
1. Click "Mail Merge" → "Send Emails with GmailApp (full features)" 
2. Enter the subject line of your Gmail draft
3. Script sends emails immediately using GmailApp with smart emoji conversion
4. Spreadsheet shows timestamp in the "Email Sent" column

## Mode Comparison

| Feature | Draft Mode | MailApp Send | GmailApp Send |
|---------|------------|--------------|---------------|
| **Review before sending** | ✅ Yes | ❌ No | ❌ No |
| **Emoji support** | ✅ Excellent (smart conversion) | ✅ Excellent (native) | ✅ Good (smart conversion) |
| **Inline images** | ✅ Full support | ❌ Converted to attachments | ✅ Full support |
| **Gmail features** | ✅ All features | ⚠️ Limited | ✅ All features |
| **From address control** | ✅ Yes | ❌ No | ✅ Yes |
| **Send-as aliases** | ✅ Yes | ❌ No | ✅ Yes |
| **Automation** | ⚠️ Manual sending required | ✅ Fully automated | ✅ Fully automated |

## Multiple Recipients

You can specify multiple email addresses in the `Recipient` column by separating them with commas:
- `user1@example.com,user2@example.com,user3@example.com`
- The first email becomes the primary recipient (TO)
- Additional emails are added as CC recipients
- All recipients receive the same personalized message

## Rich Text and Emojis

The script uses MailApp for reliable Unicode/emoji support:
- HTML formatting (bold, italic, colors, etc.)
- Unicode characters and emojis 🎉 (excellent handling with MailApp)
- File attachments
- Note: Inline images become regular attachments (MailApp limitation)

## Template Variables

Use `{{ColumnName}}` in your email draft to insert data from your spreadsheet:
- `{{Name}}` - inserts value from "Name" column
- `{{Company}}` - inserts value from "Company" column
- Any column header can be used as a variable

## Error Handling

- Failed emails will show error messages in the "Email Sent" column
- Successfully sent emails show the timestamp
- Re-running the script skips already sent emails

## Contributing

This project uses [Conventional Commits](https://www.conventionalcommits.org/) for commit messages.

Example: `feat: add support for BCC recipients`

## Why did I spend time on this?

You'd think mail merge from gmail would be totally solved problem at this point, but in 202509, I still hit issues:
* [GMail first class support](https://support.google.com/mail/answer/12921167?hl=en)
  * 👎 Required paying for Google Workspace.  I was ok to do this because could get it for 14 days free.
  * 👎 It ended up not identifying the columns in my spreadsheet correctly and insisting on a name, when I just want a column with one or more emails concatenated in a given cell if I wanted to sent to multiple people.
* Mail Merge for Gmail
  * 👎 To remove their branding from emails means a paid plan, and paid play is annual, so like $36 a year.
* Yet Another Mail Merge
  * I have used this in the past and it works great, but..
  * 👎 Free plan limited to 20 recipients per day, and I had 100 to send to.  

## Things to be aware of
* MailApp has good emoji support but limits the number of emails that can be sent for a day
* GmailApp doesn't have as good of emoji support.  See https://stackoverflow.com/questions/50686254/how-to-insert-an-emoji-into-an-email-sent-with-gmailapp for an explainer.  There is no workaround for emojis in the subject line.
* 

## Attribution

Based on the original Gmail/Sheets mail merge script from Google Apps Script samples:
- **Source**: https://developers.google.com/apps-script/samples/automations/mail-merge
- **Original Author**: Martin Hawksey
- **License**: Apache License 2.0

## License

Licensed under the Apache License, Version 2.0. See the original license header in `mail-merge.gs` for full details.

## Enhancements Made

This version includes the following improvements over the original:

1. **Three Operation Modes**: Draft creation, MailApp sending, or GmailApp sending
2. **Smart Emoji Handling**: Automatic emoji-to-HTML conversion for GmailApp compatibility  
3. **Multiple Recipients Support**: Send to multiple email addresses per row using comma-separated format
4. **Flexible Workflow Options**: Choose between review-first (drafts) or automated sending
5. **Full Gmail Feature Support**: Send-as aliases, inline images, rich formatting (draft/GmailApp modes)
6. **Enhanced Documentation**: Comprehensive setup and usage instructions with mode comparison
7. **Conventional Commits**: Standardized commit message format for better project maintenance

---

*Enhanced with ❤️ using Claude Code*