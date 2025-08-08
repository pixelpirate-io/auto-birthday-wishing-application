# Birthday WhatsApp Bot

A simple Python automation script to send personalized birthday wishes through WhatsApp using contact data from an Excel sheet.

## Features

- Reads contact information from `contacts.xlsx`
- Matches birthdays with the current date
- Sends personalized WhatsApp messages to birthday contacts
- Uses `pywhatkit` for sending messages via WhatsApp Web
- Provides real-time logs for each operation step

## Requirements

- Python 3.7+
- Google Chrome browser
- Stable internet connection (for WhatsApp Web)

## Dependencies

Install all required Python packages using pip:

```bash
pip install pandas openpyxl schedule pywhatkit
