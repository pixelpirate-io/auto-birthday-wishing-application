import pandas as pd
import schedule
import datetime
import pywhatkit
import time

# Load contacts from an Excel file
def load_contacts():
    try:
        contacts = pd.read_excel("contacts.xlsx")
        print("Contacts loaded successfully!")
        print(contacts)  # Print the loaded contacts to verify
        return contacts
    except FileNotFoundError:
        print("contacts.xlsx file not found. Please create an Excel file with 'Name', 'Phone', and 'Birthday' columns.")
        return None

# Function to check today's birthdays
def check_and_send_messages():
    print("Checking for today's birthdays...")
    contacts = load_contacts()
    if contacts is None:
        print("No contacts loaded. Exiting function.")
        return  # Exit if the file is not found
    
    today = datetime.datetime.now().strftime("%m-%d")
    print(f"Today's date: {today}")
    
    for index, contact in contacts.iterrows():
        birthday = contact['Birthday'].strftime("%m-%d")
        print(f"Checking {contact['Name']}'s birthday: {birthday}")
        if birthday == today:
            print(f"Found a match! Sending birthday message to {contact['Name']}.")
            send_whatsapp_message(contact['Phone'], contact['Name'])
        else:
            print(f"{contact['Name']}'s birthday is not today. Skipping.")

# Function to send a WhatsApp message
def send_whatsapp_message(phone, name):
    print(f"Preparing to send message to {name} at {phone}.")
    message = f"Happy Birthday, {name}! 🎉 Hope you have an amazing day filled with joy and laughter!"
    # Ensure the number has the "+" sign (if it wasn't included in Excel).
    if not str(phone).startswith('+'):
        phone = '+' + str(phone)
    print(f"Sending message to {phone}...")
    # Send message instantly via WhatsApp
    pywhatkit.sendwhatmsg_instantly(phone, message)
    print(f"Sent WhatsApp message to {name} at {phone}.")

# Schedule the task to run immediately after the script starts
def run_now():
    print("Running birthday checking and sending messages now...")
    check_and_send_messages()  # Run the task immediately

    while True:
        print("Waiting for the next execution...")
        time.sleep(1)

# Main program setup
if __name__ == "__main__":
    print("Starting the birthday bot script...")
    run_now()  # Run the task immediately
