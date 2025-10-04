# Automated Refund Email Sender (Google Forms & Sheets)

This Google Apps Script automates the process of sending personalized refund emails to students who submit a refund request via a Google Form. When a new response is recorded in the "Refund Responses" sheet, the script validates the input and sends a tailored refund confirmation email to the student, marking the status in the sheet for future reference.

---

## Features

- **Automatic Email Sending:**  
  On every new form submission, an email is sent to the respondent using the email address they provided.

- **Personalized Messaging:**  
  Each email is addressed to the student by name and gives a clear refund timeline and support information.

- **Duplicate & Error Handling:**  
  - Ensures emails are not sent multiple times.
  - Marks invalid or error entries in the sheet for easy review.

- **Self-Healing Column Tracking:**  
  If the "Email Sent" column is missing, it is automatically added.

---

## How It Works

1. **Form Submission:**  
   - The student fills out a Google Form requesting a refund.
   - The responses are saved in a Google Sheet ("Refund Responses" tab).

2. **Script Trigger:**  
   - The `onFormSubmit` function is set to run on form submit (set up as an "On form submit" trigger in Apps Script).
   - The script reads the new entry, validates the name and email, and sends an email.

3. **Sheet Update:**  
   - If an email is sent, the script marks "SENT" in the "Email Sent" column.
   - If invalid data or an error occurs, it marks "INVALID" or "ERROR" respectively.

---

## Setup

1. **Google Form & Sheet:**  
   - Set up your Google Form to collect at least the student's full name and email address.
   - Link the form to a Google Sheet (with a tab named "Refund Responses").

2. **Script Installation:**  
   - Paste the script into the Apps Script editor attached to the response Sheet.
   - Save and authorize permissions for Gmail and Sheets.

3. **Trigger Setup:**  
   - In the Apps Script editor, go to **Triggers**.
   - Add a trigger for `onFormSubmit` for the "Refund Responses" sheet, set to run on "On form submit".

---

## Customization

- **Email Template:**  
  Change the `subject` and `body` variables to customize your message.
- **Column Headers:**  
  Adjust the header-matching regex in the script if your column titles differ.
- **Support Address:**  
  Update the sign-off or support information as needed.

---

## Troubleshooting

- **Not sending emails:**  
  - Check that the trigger is set up and authorized.
  - Ensure "Refund Responses" is the correct tab name.
  - Verify that the "Full Name" and "Email Address" columns exist and are correctly titled.

- **Emails not marked:**  
  - The script adds an "Email Sent" column if missing, but ensure you allow editing of the sheet.

- **Gmail quota exceeded:**  
  - Google Apps Script has daily send limits. For large batches, throttle submissions or use a paid Google Workspace account.

---

## Example Usage

On every new refund request:
- The student receives a message like:
  ```
  Dear John Doe,

  We’ve received your refund request, and the amount will be processed and refunded within 15–20 working days.

  We appreciate your patience during this time. If you have any questions or need further assistance, feel free to reply to this email. We’re here to help.

  Best regards,
  Team abc
  ```
- The sheet marks "SENT" next to the student's row.

---

## License

MIT License
---

## Author

- Shahzadi Eman
- Shahzadieman1122@gmail.com

---
