# Email to WhatsApp Notifier

This Python script connects to your Gmail account, monitors incoming emails, and sends notifications to your WhatsApp account when certain conditions are met. It's a versatile tool that can help you stay informed about important emails.

## Prerequisites

Before using this script, you need to set up a few things:

- **Gmail Account**: You should have a Gmail account to monitor. Make sure to enable "Less secure apps" for your Gmail account.

- **WhatsApp Account**: You'll need a WhatsApp account to receive notifications.

- **Python**: Ensure that you have Python installed on your system.

- **Python Libraries**: Install the necessary Python libraries using pip:



pip install imaplib email openpyxl pywhatkit python-dotenv




## Setup

1. Clone this repository or download the script.

2. Create a `.env` file in the same directory as the script with the following contents:

GMAIL_USERNAME=your@gmail.com
GMAIL_PASSWORD=your_password
SEARCH_TEXT=your_search_text
NUMBER1=your_phone_number1
NUMBER2=your_phone_number2



Replace the placeholders with your Gmail username and password, the text you want to search for in emails (`SEARCH_TEXT`), and the phone numbers you want to send WhatsApp messages to (`NUMBER1` and `NUMBER2`).

## Usage

Run the script using:
python script.py


