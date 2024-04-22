import win32com.client
import datetime as dt
import time

import os
import pandas as pd


def check_outlook():
    email_data = {
    'Subject': [],
    'ReceivedTime': [],
    'Text': [],
    'Attachments': [],
    'Attachments_numbers': []
    
    }
    
    df = pd.DataFrame()
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 refers to the Inbox folder

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    try:

        while True:
            current_time = dt.datetime.now()
            print(f"Last email checking time was: {current_time}")
            duration_filter = current_time - dt.timedelta(minutes = 1)
            print("Duration: ", duration_filter)

            filtered_emails = messages.Restrict("[ReceivedTime] >= '" + duration_filter.strftime('%m/%d/%Y %H:%M')+ "'")
            print(len(filtered_emails))

            for msg in filtered_emails:
                print("New email arrived!")
                print("Subject:", msg.Subject)
                print("Received Time:", msg.ReceivedTime)
                print("----------")

                if "Test" in str(msg.Subject) and msg.SenderEmailAddress == "rickymehra299@gmail.com":
            
                    # Extract email subject, received time, and text
                    subject = msg.Subject.strip()
                    received_time = msg.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S") if msg.ReceivedTime is not None else None
                    text = msg.Body.strip()
                    
                    folder_path = os.path.join(os.getcwd(), "Email_Attachments_Folder")
                    
                    if not os.path.exists(folder_path):
                        os.makedirs(folder_path, exist_ok=True)
                    
                    attachments = []
                    for attch in msg.Attachments:
                        attachments.append(attch.FileName)  # Append the entire file name
                        attch.SaveAsFile(os.path.join(folder_path, attch.FileName))
                    
                    if attachments:
                        attachments_str = ', '.join(attachments)
                        attachments_counts = len(attachments)
                    else:
                        attachments_str = "No attachment"
                        attachments_counts = 0

                    # Append values to email_data dictionary
                    email_data['Subject'].append(subject)
                    email_data['ReceivedTime'].append(received_time)
                    email_data['Text'].append(text)
                    email_data['Attachments'].append(attachments_str)
                    email_data['Attachments_numbers'].append(attachments_counts) 

            # Convert the dictionary to a Pandas DataFrame
            if df.empty:
                df = pd.DataFrame(email_data)
                email_data = {key: [] for key in email_data.keys()}
                print(df.shape)
            else:
                df = pd.concat([df, pd.DataFrame(email_data)], ignore_index=True)
                df.drop_duplicates(ikeep='first', nplace=True)
                email_data = {key: [] for key in email_data.keys()}
                print(df.shape)

            print(df)

            # Wait for 2 minutes before checking again
            new_time = dt.datetime.now()
            execuation_time = new_time - current_time
            print(execuation_time)

            remaining_sleep_time = 60-execuation_time.total_seconds()
            print(remaining_sleep_time)
            time.sleep(remaining_sleep_time)

    except KeyboardInterrupt:
        # Save the DataFrame to an Excel file when user stops execution
        #df.drop_duplicates(inplace=True)
        df.to_excel("outlook_emails.xlsx", index=False)
        print("Execution stopped. DataFrame saved to 'outlook_emails.xlsx'.")


if __name__ == "__main__":
    df = check_outlook()
    

