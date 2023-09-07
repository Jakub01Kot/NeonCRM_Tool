# Import required modules
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime, timedelta
import pytz
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
import os


load_dotenv()

# Function to read a DataFrame from an Excel file
def read_dataframe_from_excel(file_path):
    return pd.read_excel(file_path)


# Function to send email
def send_email(subject, body, to_email):

    # Email authentication and setup
    # Chamber's Email
    from_email = os.getenv("NOTIFICATION_EMAIL_FOR_NEW_SIGNUPS")
    #from_email = "chambernotification@gmail.com"

    # App-Password from Gmail 2-step authentication
    #password = "loxqguewogwqjzte"
    password = os.getenv("2_STEP_AUTHENTICATION_PW_FOR_EMAIL_ACCOUNT")
    toaddr = to_email
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = toaddr
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Connect to email server and send email
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(msg['From'], password)
    server.sendmail(from_email, toaddr, msg.as_string())
    server.quit()


# Function to check if a contact exists in NeonCRM
def check_contact_exists(org_id, api_key, first_name, last_name):

    # API endpoint and headers
    url = "https://api.neoncrm.com/v2/accounts/search"
    auth = HTTPBasicAuth(org_id, api_key)
    headers = {"Content-Type": "application/json", "NEON-API-VERSION": "2.6"}

    # API payload for search
    payload = {
        "searchFields": [
            {"field": "First Name", "operator": "EQUAL", "value": first_name},
            {"field": "Last Name", "operator": "EQUAL", "value": last_name}
        ],
        "outputFields": ["Last Name"],
        "pagination": {"currentPage": 0, "pageSize": 1}
    }

    # Make the API request
    response = requests.post(url, headers=headers, json=payload, auth=auth)

    # Check if the contact exists
    if response.status_code == 200:
        json_response = response.json()
        total_results = json_response['pagination']['totalResults']
        if total_results > 0:
            return True
    return False


# Function to create a new activity in NeonCRM
def create_activity(org_id, api_key, client_id, full_name, linkedin_url, swiss_connection, experience_in_months, system_user_name):

    # API endpoint and headers
    url = "https://api.neoncrm.com/v2/activities"
    headers = {"Content-Type": "application/json", "NEON-API-VERSION": "2.6"}

    # Timezone and current date-time setup
    tz = pytz.timezone('UTC')
    now = datetime.now(tz)
    one_hour_later = now + timedelta(hours=1)
    # API payload for creating a new activity
    payload = {
        "subject": "Welcome Email",
        "note": "A welcome email will be sent to the new contact.",
        "activityDates": {
            "startDate": now.date().isoformat(),
            "endDate": one_hour_later.date().isoformat(),
            "timeZone": {
                "id": "4",
                "name": "UTC",
                "status": "ACTIVE"
            }
        },
        "clientId": client_id,
        "systemUserId": "12340",
        "status": {
            "id": "3",
            "name": "Open",
        },
        "priority": "High"
    }
    # Make the API request to create the activity
    auth = HTTPBasicAuth(org_id, api_key)
    response = requests.post(url, headers=headers, json=payload, auth=auth)

    # Check for successful activity creation
    if response.status_code == 200:
        neoncrm_url = f"https://saccsf.app.neoncrm.com/admin/accounts/{client_id}/about?nlListId=5154615&nlIndex=0&nlCache=accountId&lastQuery=true"

        # Convert experience in months to years and months
        years = experience_in_months // 12
        months = experience_in_months % 12

        # Change the message's content based on the contact's experience
        if years == 0:
            experience_message = f"you worked at {swiss_connection} for {months} months."
        else:
            experience_message = f"you worked at {swiss_connection} for {years} years and {months} months."

        # Short version of the message
        short_message = f"""Hi {full_name},
        
I found your profile and noticed that {experience_message} This Swiss connection makes you an interesting candidate for the Swiss American Chamber of Commerce San Francisco. The chamber provides exclusive events, business insights, and a diverse network of professionals eager to explore opportunities with someone of your background.

We would be pleased to welcome you to SACCSF and build a network together. Feel free to reach out if you have any questions and would like to discuss this further. For more information, you can visit our website: www.saccsf.com.

We are looking forward to hearing from you.

Kind regards,
{system_user_name}"""

        # Long version of the message
        long_message = f"""Dear {full_name},

I hope this message finds you well. I recently came across your profile and was particularly interested in your experience at {swiss_connection}. It seems to me that you could greatly benefit from joining the Swiss American Chamber of Commerce San Francisco (SACCSF).

Our chamber is a non-profit organization that facilitates connections, collaboration, and growth between Switzerland and the United States. Whether you have personal ties to Switzerland or are involved in Swiss-related business, SACCSF offers a unique platform for networking and expanding your horizons.

As a member, you will gain access to exclusive events, business insights, and a diverse network of professionals eager to explore opportunities with someone of your background. We would be pleased to welcome you to SACCSF and build a network together. Feel free to reach out if you have any questions or would like to discuss this further. You can also visit our website for more information.

We look forward to hearing from you.

Kind regards,
{system_user_name}"""

        email_body = f"A new activity has been created for client {client_id}.\n"

        email_body += f"Full Name: {full_name}\n"
        email_body += f"LinkedIn URL: {linkedin_url}\n\n"  # Adding two line breaks here
        email_body += f"NeonCRM Profile: {neoncrm_url}\n\n"  # Adding two line breaks here

        # Short Message
        email_body += "Short Message:\n"  # Adding a line break here
        email_body += short_message.strip()  # Removing leading and trailing spaces
        email_body += "\n\n"  # Adding two line breaks here

        # Long Message
        email_body += "Long Message:\n"  # Adding a line break here
        email_body += long_message.strip()  # Removing leading and trailing spaces
        email_body += "\n"  # Adding a line break here

        # Define recipient email address
        to_email = os.getenv("NOTIFICATION_RECIPIENT_TEMPLATE_FOR_CONTACTING")
        # to_email = "jakub.kot@stud.hslu.ch"

        # Function to send an email upon successful activity creation
        send_email("New Activity Created", email_body, to_email)

        print(f"Successfully sent an email to {to_email}.")

    else:
        print("Failed to send an email.")


# Function to add a new contact to NeonCRM
def add_contact_to_neoncrm(org_id, api_key, full_name, linkedin_url, swiss_connection, experience_in_months):

    # Split the full name into first and last names
    first_name, last_name = full_name.split(" ", 1)

    # Check if the contact already exists in NeonCRM
    if check_contact_exists(org_id, api_key, first_name, last_name):
        print(f"Contact with name {first_name} {last_name} already exists.")
        return

    # Define the NeonCRM API URL and headers
    url = "https://api.neoncrm.com/v2/accounts"
    auth = HTTPBasicAuth(org_id, api_key)
    headers = {"Content-Type": "application/json", "NEON-API-VERSION": "2.6"}

    # Define the payload for creating a new individual account
    payload = {
        "individualAccount": {
            "primaryContact": {
                "firstName": first_name,
                "lastName": last_name
            }
        }
    }

    # Perform a POST request to create a new contact in NeonCRM
    response = requests.post(url, headers=headers, json=payload, auth=auth)

    # Check the response status code
    if response.status_code == 200:
        client_id = response.json().get('id', '')
        print(f"Successfully created a new contact with name {full_name} and user ID {client_id}.")

        # Defining the name to be shown on the messages
        system_user_name = os.getenv("NAME_TO_BE_SHOWN_ON_TEMPLATES_AS_SENDER")
        #system_user_name = "Marjorie"

        create_activity(org_id, api_key, client_id, full_name, linkedin_url, swiss_connection, experience_in_months,
                        system_user_name)
    else:
        print("Failed to create a new contact.")


# Main execution starts here
if __name__ == "__main__":
    # org_id = 'saccsf'  # Define organization ID
    org_id = os.getenv("ORG_ID")
    api_key = os.getenv("API_KEY")
    # api_key = 'db373b2ba6403dd890083126f68e10a8'  # Define API key

    # Path to existing Excel file
    existing_file_path = '/Users/jakub/Downloads/existing_spreadsheet.xlsx'

    # Read the Excel file and store the data in unique_rows
    unique_rows = read_dataframe_from_excel(existing_file_path)

    # Loop through each row in unique_rows DataFrame
    for index, row in unique_rows.iterrows():
        full_name = row['Full Name']  # Extract Full Name from row
        linkedin_url = row['LinkedIn (url)']  # Extract LinkedIn URL from row
        swiss_connection = row['Swiss Connection']  # Extract Company from row
        experience_in_months = row['Experience']  # Extract Experience from row

        # Add contact to NeonCRM
        add_contact_to_neoncrm(org_id, api_key, full_name, linkedin_url, swiss_connection, experience_in_months)
