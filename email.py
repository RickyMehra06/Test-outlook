import win32com.client
import datetime as dt
import time

import pandas as pd

def check_outlook():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 refers to the Inbox folder

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    while True:
        current_time = dt.datetime.now()
        print(f"Last email checking time was: {current_time}")
        duration_filter = current_time - dt.timedelta(minutes = 120)
        print("Duration: ", duration_filter)

        filtered_emails = messages.Restrict("[ReceivedTime] >= '" + duration_filter.strftime('%m/%d/%Y %H:%M')+ "'")
        print(len(filtered_emails))

        for message in filtered_emails:
            print("New email arrived!")
            print("Subject:", message.Subject)
            print("Received Time:", message.ReceivedTime)
            print("----------")

        # Wait for 2 minutes before checking again
        new_time = dt.datetime.now()
        execuation_time = new_time - current_time
        print(execuation_time)

        remaining_sleep_time = 120-execuation_time.total_seconds()
        print(remaining_sleep_time)
        time.sleep(remaining_sleep_time)


if __name__ == "__main__":
    check_outlook()
