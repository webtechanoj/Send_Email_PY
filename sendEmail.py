"""
This program is designed to work for Office 365 hosted emails
REQUIREMENTS -
1. Contacts.txt on the same directory containing Name and Email
2. Message.txt with the message you want to send across


"""

import smtplib

from string import Template

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

MY_ADDRESS = input("Enter your email: ")
PASSWORD = input("Enter your password: ")
def get_contacts(filename):
    """
    Return two lists names, emails containing names and email addresses
    read from a file specified by filename.
    """
    
    names = []
    emails = []
    with open(filename, mode='r', encoding='utf-8') as contacts_file:
        for a_contact in contacts_file:
            names.append(a_contact.split()[0])
            emails.append(a_contact.split()[1])
    return names, emails

def read_template(filename):
    """
    Returns a Template object comprising the contents of the 
    file specified by filename.
    """
    
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

def main():
    names, emails = get_contacts('Contacts.txt') # read contacts 
    message_template = read_template('Message.txt')
    try:
        # Assign the SMTP server settings here if they are different from Office365
        s = smtplib.SMTP(host='smtp.office365.com', port='587')
        s.starttls()
        s.login(MY_ADDRESS, PASSWORD)
        
        # For each contact, send the email:
        for name, email in zip(names, emails):
            msg = MIMEMultipart()       # create a message

            # add in the actual person name to the message template
            message = message_template.substitute(PERSON_NAME=name.title())

            # Prints out the message body for our sake
            print(message)

            # setup the parameters of the message
            msg['From']=MY_ADDRESS
            msg['To']=email

            msg['Subject']="This is TEST"
            print("The email has been sent")
            # add in the message body
            msg.attach(MIMEText(message, 'plain'))
            
            # send the message via the server set up earlier.
            s.send_message(msg)
            del msg
    
    except:
        print("There was an error while sending the email")
    
    # Terminate the SMTP session and close the connection
    s.quit()
    
if __name__ == '__main__':
    main()