# YAMM-Like Mail Merge Web App (Google Apps Script)

A modern, user-friendly mail merge web app for Google Sheets and Gmail, featuring:
- Rich text editor with font & size selection, color, highlight, code, blockquote, image and link insertion (including masked text for links)
- Sender display name masking (`From: Name <email@example.com>`) if supported by Gmail
- Google Sheets integration: select sheet, recipient column, range, and column-to-placeholder mapping
- Live preview of emails as you edit
- Easy-to-use UI, styled after Gmail compose

---

## Features

- **Rich Email Composer**: Format text, insert images by URL, insert/edit links with masked text, choose font & size, colors, and more.
- **Personalization**: Use `{{placeholders}}` in subject/body and map them to columns from your Google Sheet.
- **Sender Masking**: Optionally specify a sender name for the "From" field.
- **Sheet Picker**: Enter a sheet URL or ID, select worksheet, email column, data range, and map columns to placeholders.
- **Testing**: See exactly how your email will look by sending email to your own inbox
- **Easy Deployment**: No external libraries needed, 100% Apps Script + Quill.js.
---

## Setup

### 1. Copy Files
- Import `YAMM_Like_Mailer_UI.html` and `Code.gs` into a new Google Apps Script project.

### 2. Deploy as Web App
- In the Apps Script editor, go to **Deploy > New deployment**
- Select **Web app**
- Set "Execute as" to **Me** (the owner)
- Set "Who has access" to **Anyone with the link** (or restrict as you wish)
- Deploy and copy the web app URL

### 3. Permissions
- The first time you run, you'll be asked to authorize access to Gmail and Sheets.

---

## Usage

1. Open the web app.
2. Enter your Google Sheet's URL or ID and load.
3. Select the worksheet, recipient column, and range.
4. Map placeholders to columns as needed.
5. Compose your email:
   - Use the toolbar for formatting, font, size, color, highlight, code, quote, etc.
   - Insert images via URL with the image button.
   - Insert/edit links with custom text mask.
6. Optionally, set a sender display name in the "From" box (see below).
7. Click **Send** to mail-merge to all rows in the range.

---

## Notes on Sender Masking ("From" Name)

- You can set a sender name (`From: Name <email@example.com>`) in the UI.

---

## Known Limitations

- If you schedule multiple email batches for the exact same time, only one will be sent. This is due to how the backend stores scheduled jobs by timestamp. To ensure all scheduled emails are sent, use unique times for each batch.
- Quota reports from Google are not correct.
- Attachment feature is currently challenging to implement, so it is not supported at this time.

---

## FAQ

**Can I attach files?**  
No, this version supports image insertion by URL, not file attachments.

**Can I use external fonts?**  
Font selection includes several popular web-safe fonts (Sans Serif, Serif, Monospace, Roboto, Arial, Times New Roman, Courier New).

**Can I use this for mass mail?**  
Yes, but respect Gmail's daily sending limits and anti-spam policies.

---

## Credits

- UI inspired by Gmail
- Rich text editing powered by [Quill.js](https://quilljs.com/)

---

## License

MIT License (see [LICENSE](LICENSE))
