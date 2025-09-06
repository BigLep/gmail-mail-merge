# Gmail/Sheets Mail Merge

Enhanced mail merge script for Google Apps Script that sends personalized emails using Gmail drafts as templates and Google Sheets for recipient data.

## Features

- Use Gmail draft messages as email templates
- Support for rich text content including emojis and HTML formatting
- Multiple recipients per row (comma-separated format) 
- Template variable substitution using `{{variable}}` syntax
- File attachments support
- Email tracking with sent status column
- Error handling and logging
- Uses MailApp for reliable emoji/Unicode handling

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

### 3. Run Mail Merge
1. In Google Sheets, you'll see a "Mail Merge" menu appear
2. Click "Mail Merge" → "Send Emails"
3. Enter the subject line of your Gmail draft
4. The script will send personalized emails and update the "Email Sent" column

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

This project uses [Conventional Commits](https://www.conventionalcommits.org/) for commit messages:

- `feat:` - New features
- `fix:` - Bug fixes
- `docs:` - Documentation changes
- `style:` - Code style changes
- `refactor:` - Code refactoring
- `test:` - Adding or updating tests
- `chore:` - Maintenance tasks

Example: `feat: add support for BCC recipients`

## Attribution

Based on the original Gmail/Sheets mail merge script from Google Apps Script samples:
- **Source**: https://developers.google.com/apps-script/samples/automations/mail-merge
- **Original Author**: Martin Hawksey
- **License**: Apache License 2.0

## License

Licensed under the Apache License, Version 2.0. See the original license header in `mail-merge.gs` for full details.

## Enhancements Made

This version includes the following improvements over the original:

1. **Multiple Recipients Support**: Send to multiple email addresses per row using comma-separated format
2. **Reliable Emoji Support**: Uses MailApp instead of GmailApp for proper Unicode/emoji encoding
3. **Enhanced Documentation**: Comprehensive setup and usage instructions
4. **Conventional Commits**: Standardized commit message format for better project maintenance

---

*Enhanced with ❤️ using Claude Code*