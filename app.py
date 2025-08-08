from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd
import datetime
import pywhatkit
import time
import os

app = Flask(__name__)
app.config['SECRET_KEY'] = 'simple'
# Configuration
EXCEL_FILE = 'contacts.xlsx'

def load_contacts():
    """Load contacts from Excel file"""
    try:
        if os.path.exists(EXCEL_FILE):
            contacts = pd.read_excel(EXCEL_FILE)
            # Ensure the Birthday column is datetime
            contacts['Birthday'] = pd.to_datetime(contacts['Birthday'])
            print(f"Loaded {len(contacts)} contacts from {EXCEL_FILE}")
            return contacts
        else:
            # Create default contacts with today's birthday for demo
            print(f"{EXCEL_FILE} not found. Creating with default data...")
            default_contacts = pd.DataFrame({
                'Name': ['Alice Johnson', 'Today Birthday Person', 'Bob Smith'],
                'Phone': ['+1234567890', '+9876543210', '+5555555555'],
                'Birthday': [
                    pd.to_datetime('1995-03-15'),
                    pd.to_datetime('2000-08-08'),  # Today's date for demo
                    pd.to_datetime('1988-12-25')
                ]
            })
            # Save the default contacts
            if save_contacts(default_contacts):
                print("Default contacts created successfully!")
                return default_contacts
            else:
                return pd.DataFrame(columns=['Name', 'Phone', 'Birthday'])
    except Exception as e:
        print(f"Error loading contacts: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame(columns=['Name', 'Phone', 'Birthday'])

def save_contacts(contacts_df):
    """Save contacts to Excel file"""
    try:
        contacts_df.to_excel(EXCEL_FILE, index=False)
        print(f"Contacts saved successfully to {EXCEL_FILE}")
        return True
    except Exception as e:
        print(f"Error saving contacts: {e}")
        import traceback
        traceback.print_exc()
        return False

def validate_contact(name, phone, birthday):
    """Validate contact data"""
    errors = []
    
    if not name or name.strip() == '':
        errors.append("Name is required")
    
    if not phone or phone.strip() == '':
        errors.append("Phone number is required")
    
    if not birthday:
        errors.append("Birthday is required")
    else:
        try:
            datetime.datetime.strptime(birthday, '%Y-%m-%d')
        except ValueError:
            errors.append("Invalid birthday format")
    
    return errors

@app.route('/')
def index():
    """Main page showing contact list"""
    contacts = load_contacts()
    contact_list = []
    
    for index, contact in contacts.iterrows():
        contact_dict = {
            'index': index,
            'name': contact['Name'],
            'phone': contact['Phone'],
            'birthday': contact['Birthday'].strftime('%Y-%m-%d') if pd.notna(contact['Birthday']) else '',
            'birthday_display': contact['Birthday'].strftime('%B %d, %Y') if pd.notna(contact['Birthday']) else ''
        }
        contact_list.append(contact_dict)
    
    return render_template('index.html', contacts=contact_list)

@app.route('/add_contact', methods=['GET', 'POST'])
def add_contact():
    """Add new contact"""
    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        phone = request.form.get('phone', '').strip()
        birthday = request.form.get('birthday', '').strip()
        
        # Validate input
        errors = validate_contact(name, phone, birthday)
        
        if errors:
            for error in errors:
                flash(error, 'error')
            return render_template('add_contact.html', name=name, phone=phone, birthday=birthday)
        
        # Load existing contacts
        contacts = load_contacts()
        
        # Create new contact
        new_contact = pd.DataFrame({
            'Name': [name],
            'Phone': [phone],
            'Birthday': [pd.to_datetime(birthday)]
        })
        
        # Add to existing contacts
        contacts = pd.concat([contacts, new_contact], ignore_index=True)
        
        # Save to file
        if save_contacts(contacts):
            flash('Contact added successfully!', 'success')
            return redirect(url_for('index'))
        else:
            flash('Error saving contact', 'error')
    
    return render_template('add_contact.html')

@app.route('/edit_contact/<int:contact_id>', methods=['GET', 'POST'])
def edit_contact(contact_id):
    """Edit existing contact"""
    contacts = load_contacts()
    
    if contact_id >= len(contacts):
        flash('Contact not found', 'error')
        return redirect(url_for('index'))
    
    contact = contacts.iloc[contact_id]
    
    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        phone = request.form.get('phone', '').strip()
        birthday = request.form.get('birthday', '').strip()
        
        # Validate input
        errors = validate_contact(name, phone, birthday)
        
        if errors:
            for error in errors:
                flash(error, 'error')
            return render_template('edit_contact.html', contact={
                'index': contact_id,
                'name': name,
                'phone': phone,
                'birthday': birthday
            })
        
        # Update contact
        contacts.iloc[contact_id, contacts.columns.get_loc('Name')] = name
        contacts.iloc[contact_id, contacts.columns.get_loc('Phone')] = phone
        contacts.iloc[contact_id, contacts.columns.get_loc('Birthday')] = pd.to_datetime(birthday)
        
        # Save to file
        if save_contacts(contacts):
            flash('Contact updated successfully!', 'success')
            return redirect(url_for('index'))
        else:
            flash('Error updating contact', 'error')
    
    contact_data = {
        'index': contact_id,
        'name': contact['Name'],
        'phone': contact['Phone'],
        'birthday': contact['Birthday'].strftime('%Y-%m-%d') if pd.notna(contact['Birthday']) else ''
    }
    
    return render_template('edit_contact.html', contact=contact_data)

@app.route('/delete_contact/<int:contact_id>', methods=['POST'])
def delete_contact(contact_id):
    """Delete contact"""
    contacts = load_contacts()
    
    if contact_id >= len(contacts):
        flash('Contact not found', 'error')
        return redirect(url_for('index'))
    
    # Remove contact
    contacts = contacts.drop(contacts.index[contact_id]).reset_index(drop=True)
    
    # Save to file
    if save_contacts(contacts):
        flash('Contact deleted successfully!', 'success')
    else:
        flash('Error deleting contact', 'error')
    
    return redirect(url_for('index'))

@app.route('/check_birthdays')
def check_birthdays():
    """Check for today's birthdays and send messages"""
    contacts = load_contacts()
    today = datetime.datetime.now().strftime("%m-%d")
    birthday_contacts = []
    
    for index, contact in contacts.iterrows():
        if pd.notna(contact['Birthday']):
            birthday = contact['Birthday'].strftime("%m-%d")
            if birthday == today:
                birthday_contacts.append({
                    'name': contact['Name'],
                    'phone': contact['Phone']
                })
    
    return render_template('birthday_check.html', 
                         birthday_contacts=birthday_contacts, 
                         today=datetime.datetime.now().strftime("%B %d, %Y"))

@app.route('/send_birthday_messages', methods=['POST'])
def send_birthday_messages():
    """Send birthday messages to today's birthday contacts"""
    try:
        contacts = load_contacts()
        today = datetime.datetime.now().strftime("%m-%d")
        sent_count = 0
        
        for index, contact in contacts.iterrows():
            if pd.notna(contact['Birthday']):
                birthday = contact['Birthday'].strftime("%m-%d")
                if birthday == today:
                    # Send WhatsApp message
                    phone = str(contact['Phone'])
                    if not phone.startswith('+'):
                        phone = '+' + phone
                    
                    message = f"Happy Birthday, {contact['Name']}! 🎉 Hope you have an amazing day filled with joy and laughter!"
                    
                    try:
                        pywhatkit.sendwhatmsg_instantly(phone, message)
                        sent_count += 1
                        time.sleep(2)  # Small delay between messages
                    except Exception as e:
                        flash(f"Error sending message to {contact['Name']}: {str(e)}", 'error')
        
        if sent_count > 0:
            flash(f'Successfully sent {sent_count} birthday messages!', 'success')
        else:
            flash('No birthday messages to send today.', 'info')
            
    except Exception as e:
        flash(f'Error checking birthdays: {str(e)}', 'error')
    
    return redirect(url_for('check_birthdays'))

if __name__ == '__main__':
    app.run(debug=True, host='127.0.0.1', port=5000)