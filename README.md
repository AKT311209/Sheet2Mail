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

- Import `YAMM_Like_Mailer_UI.html` and `Code.gs` into a new Google Apps Script project.

### 2. Deploy as Web App
- Select **Web app**
- Set "Who has access" to **Anyone with the link** (or restrict as you wish)
- Deploy and copy the web app URL

---
## Usage

For detailed usage instructions and step-by-step guidance, please see the [User Guide](./guide.md).

## Notes on Sender Masking ("From" Name)


## Known Limitations

- If you schedule multiple email batches for the exact same time, only one will be sent. This is due to how the backend stores scheduled jobs by timestamp. To ensure all scheduled emails are sent, use unique times for each batch.
- Quota reports from Google are not correct.

---
For detailed usage instructions and step-by-step guidance, please see the [User Guide](./guide.md).
...existing code...

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
