import smtplib
import pandas as pd
from datetime import datetime
from email.message import EmailMessage
import os

def load_namedays(filename):
    nameday_dict = {}
    with open(filename, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = line.split(" ", 1)
            if len(parts) < 2:
                continue
            date, names = parts
            for name in names.replace(" a ", "/").replace(",", "/").split("/"):
                name = name.strip().lower()
                if name:
                    nameday_dict[name] = date
    return nameday_dict

df = pd.read_excel("birthday.xlsx", index_col=[0])
nameday_dict = load_namedays("name_days.txt")
now = datetime.now()

def fill_null_cells(where):
    for i,item in enumerate(df[where]):
        if pd.isna(item):
            df.loc[i+1, where] = "none"

fill_null_cells("week_prior_sent")
fill_null_cells("exact_day_sent")

for i,item in enumerate(df["birthday_date"]):
    if pd.isna(item):
        if pd.notna(df.loc[i+1, "birthday_full"]):
            temp = str(df.loc[i+1, "birthday_full"])
            temp = datetime.strptime(temp,"%Y-%m-%d %H:%M:%S")
            df.loc[i+1, "birthday_date"] = datetime.strftime(temp, "%d.%m.")
        else:
            continue

for i,item in enumerate(df["nameday"]):
    if not pd.notna(item):
        if str(df.loc[i+1,"name"]).lower() in nameday_dict.keys():
            nameday_date = nameday_dict[df.loc[i+1,"name"].lower()]
            df.loc[i+1, "nameday"] = nameday_date + "."
        else:
            print(f"{df.loc[i+1, 'name']} is not in nameday_dict.")
            continue

if now.year - df.loc[1,"this_year"] > 0:
    df.loc[1,"this_year"] = now.year
    for i in range(df.index.max()):
        df.loc[i+1, "week_prior_sent"] = "none"
        df.loc[i+1, "exact_day_sent"] = "none"
df.to_excel("birthday.xlsx")

def send_mail(subject,text,receiver):
    email = os.environ.get("EMAIL_ADDRESS")
    password = os.environ.get("EMAIL_PASSWORD")
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls() #Security TLS
    s.login(email, password)

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = email
    msg["To"] = receiver
    msg.set_content(text)
    s.send_message(msg)
    s.quit()

receivers = [os.environ.get("EMAIL_ADDRESS_TO_1"),os.environ.get("EMAIL_ADDRESS_TO_2"), os.environ.get("EMAIL_ADDRESS_TO_3")]

for i,item in enumerate(df["birthday_date"]):
    if pd.notna(item):
        full_date_str = f"{item}{now.year}"
        obj = datetime.strptime((full_date_str),f"%d.%m.%Y")
        birthday = datetime(now.year,obj.month,obj.day)
        difference = now - birthday
        alias = df.loc[i+1, "alias"]
        date = df.loc[i+1, "birthday_date"]
        temp = str(df.loc[i+1, "birthday_full"])
        temp = datetime.strptime(temp,"%Y-%m-%d %H:%M:%S")
        age = now.year - temp.year
        week_prior_sent = df.loc[i+1, "week_prior_sent"]
        exact_day_sent = df.loc[i+1, "exact_day_sent"]
        if (difference.days < 0 and difference.days >= -7) and (week_prior_sent == "none" or week_prior_sent == "nday"):
            for receiver in receivers:
                send_mail(f"{alias} narozeniny!", f"{alias} bude mít dne {date} své {age}. narozeniny. Jdi koupit dárek!",receiver)
            if week_prior_sent == "none":
                df.loc[i+1,"week_prior_sent"] = "bday"
            else:
                df.loc[i+1,"week_prior_sent"] = "both"
        elif difference.days == 0 and (exact_day_sent == "none" or exact_day_sent == "nday"):
            for receiver in receivers:
                send_mail(f"{alias} narozeniny DNES!!!", f"{alias} má dnes {date} své {age}. narozeniny. Napiš jim!",receiver)
            if exact_day_sent == "none":
                df.loc[i+1,"exact_day_sent"] = "bday"
            else:
                df.loc[i+1,"exact_day_sent"] = "both"

for i,item in enumerate(df["nameday"]):
    if pd.notna(item):
        full_date_str = f"{item}{now.year}"
        obj = datetime.strptime((full_date_str),f"%d.%m.%Y")
        nameday = datetime(now.year,obj.month,obj.day)
        difference_nday = now - nameday
        alias = df.loc[i+1, "alias"]
        date = df.loc[i+1, "nameday"]
        week_prior_sent = df.loc[i+1, "week_prior_sent"]
        exact_day_sent = df.loc[i+1, "exact_day_sent"]
        if (difference_nday.days < 0 and difference_nday.days >= -7) and (week_prior_sent == "none" or week_prior_sent == "bday"):
            for receiver in receivers:
                send_mail(f"{alias} svátek!", f"{alias} bude mít dne {date} svátek. Jdi koupit dárek!",receiver)
            if week_prior_sent == "none":
                df.loc[i+1,"week_prior_sent"] = "nday"
            else:
                df.loc[i+1,"week_prior_sent"] = "both"
        elif difference_nday.days == 0 and (exact_day_sent == "none" or exact_day_sent == "bday"):
            for receiver in receivers:
                send_mail(f"{alias} svátek!!!", f"{alias} má dnes {date} svátek. Napiš jim!",receiver)
            if exact_day_sent == "none":
                df.loc[i+1,"exact_day_sent"] = "nday"
            else:
                df.loc[i+1,"exact_day_sent"] = "both"

df.to_excel("birthday.xlsx")

#TODO: generalise and post on github as public project