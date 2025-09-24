from dotenv import load_dotenv
load_dotenv(r"C:\Users\Drako\Desktop\Coding\Projects\Birthday reminder\birthday_notifier.env")
import os
email = os.environ.get("EMAIL_ADDRESS")
password = os.environ.get("EMAIL_PASSWORD")
receiver_1 = os.environ.get("EMAIL_ADDRESS_TO_1")
receiver_2 = os.environ.get("EMAIL_ADDRESS_TO_2")

receivers = os.getenv("receivers").split(',')
print(receivers)
