# PowerShell Script for Email Automation

This is a PowerShell script designed to automate the process of sending emails with attached PDF files from a selected directory. The script provides a simple GUI for inputting the email subject and body, and supports line breaks in the body text which are converted into HTML format for the email. 

## Prerequisites

- Windows system with PowerShell installed (v3.0 or higher recommended).
- Microsoft Outlook installed and set up with an active email account.

## Usage

1. Clone the repository or download the .ps1 script file.
2. Run the script in PowerShell.
3. In the GUI, fill in the 'Subject:' field with your desired email subject line.
4. In the 'Body:' field, enter the body of your email. If you want to create a new line, press `Enter`.
5. After you have entered your subject and body, click 'OK'.
6. You'll be asked to select a folder. Navigate to the directory containing the PDF files you wish to email and click 'OK'.
7. The script will then automatically send an email to recipients based on the filename of each PDF in the directory (assumes filename format is `username.pdf`, and the email address format is `username@domain.com`).

## Note

This script does not support rich text or HTML directly in the GUI, but you can add line breaks in the 'Body:' field which will be formatted correctly in the sent email.

This script is meant as a basic template and may need modifications to suit specific use cases.
