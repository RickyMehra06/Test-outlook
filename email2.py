import win32com.client
import datetime as dt
import time

import os
import pandas as pd
import mysql.connector
import mysql.connector as connection


def push_data_to_dataframe(df, email_data):

    # Convert the dictionary to a Pandas DataFrame
    if df.empty:
        df = pd.DataFrame(email_data)
        email_data = {key: [] for key in email_data.keys()}
        print(df.shape)
    else:
        df = pd.concat([df, pd.DataFrame(email_data)], ignore_index=True)
        df.drop_duplicates(keep='first', nplace=True)
        email_data = {key: [] for key in email_data.keys()}
        print(df.shape)

    return df

def push_data_to_mysql(df):

    # Connect to MySQL
    mysql_connection = mysql.connector.connect(
        host="localhost",
        user="root",
        password="12345",
        database="food_database"
    )
    mysql_cursor = mysql_connection.cursor()

    table_name = 'emails1'
    df.to_sql(name=table_name, con=mysql_cursor, if_exists='replace', index=False)
    print("Uploaded into Mysql {df.shape[0]}")

    # Close the connection
    mysql_cursor.dispose()





def main():    
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
            duration_filter = current_time - dt.timedelta(minutes = 2800)
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


            data = push_data_to_dataframe(df, email_data)

            # Wait for 2 minutes before checking again
            new_time = dt.datetime.now()
            execuation_time = new_time - current_time
            print(execuation_time)

            remaining_sleep_time = 15-execuation_time.total_seconds()
            print(remaining_sleep_time)
            time.sleep(remaining_sleep_time)

    except KeyboardInterrupt:
        # Save the DataFrame to an Excel file when user stops execution
        if data.empty:
            print("No data extracted")
        else:
            data.to_excel("outlook_emails.xlsx", index=False)
            print("DataFrame saved to 'outlook_emails.xlsx'.")
            push_data_to_mysql(data)


if __name__ == "__main__":
    main()
    

