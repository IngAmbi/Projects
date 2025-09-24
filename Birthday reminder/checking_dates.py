from datetime import datetime,timedelta
import pandas as pd

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

df = pd.read_excel(r"C:\Users\Drako\Desktop\Coding\Projects\Birthday reminder\birthday.xlsx", index_col=[0])
nameday_dict = load_namedays(r"C:\Users\Drako\Desktop\Coding\Projects\Birthday reminder\name_days.txt")

for i,item in enumerate(df["nameday"]):
    if not pd.notna(item):
        if df.loc[i+1,"name"].lower() in nameday_dict.keys():
            nameday_date = nameday_dict[df.loc[i+1,"name"].lower()]
            df.loc[i+1, "nameday"] = nameday_date + "."
        else:
            print(f"{df.loc[i+1, 'name']} is not in nameday_dict.")
            continue


now = datetime.now()
for i,item in enumerate(df["birthday_date"]):
    if pd.notna(item):
        obj = datetime.strptime(str(item),"%d.%m.")
        birthday = datetime(now.year,obj.month,obj.day)
        difference = now - birthday
        name = df.loc[i+1, "name"]
        date = df.loc[i+1, "birthday_date"]
        temp = str(df.loc[i+1, "birthday_full"])
        temp = datetime.strptime(temp,"%Y-%m-%d %H:%M:%S")
        age = now.year - temp.year
        week_prior_sent = df.loc[i+1, "week_prior_sent"]
        exact_day_sent = df.loc[i+1, "exact_day_sent"]
        if (difference.days < 0 and difference.days >= -7) and (week_prior_sent == "none" or week_prior_sent == "nday"):
            print("Email sent for birthday reminder week prior!")
            if week_prior_sent == "none":
                df.loc[i+1,"week_prior_sent"] = "bday"
            else:
                df.loc[i+1,"week_prior_sent"] = "both"
        elif difference.days == 0 and (exact_day_sent == "none" or exact_day_sent == "nday"):
            print("Email sent for birthday exact day!")
            if exact_day_sent == "none":
                df.loc[i+1,"exact_day_sent"] = "bday"
            else:
                df.loc[i+1,"exact_day_sent"] = "both"

for i,item in enumerate(df["nameday"]):
    if pd.notna(item):
        obj = datetime.strptime(str(item), "%d.%m.")
        nameday = datetime(now.year,obj.month,obj.day)
        difference_nday = now - nameday
        name = df.loc[i+1, "name"]
        date = df.loc[i+1, "nameday"]
        week_prior_sent = df.loc[i+1, "week_prior_sent"]
        exact_day_sent = df.loc[i+1, "exact_day_sent"]
        if (difference_nday.days < 0 and difference_nday.days >= -7) and (week_prior_sent == "none" or week_prior_sent == "bday"):
            print("Email sent for nameday week prior!")
            if week_prior_sent == "none":
                df.loc[i+1,"week_prior_sent"] = "nday"
            else:
                df.loc[i+1,"week_prior_sent"] = "both"
        elif difference_nday.days == 0 and (exact_day_sent.days == "none" or exact_day_sent == "bday"):
            print("Email sent for nameday exact day!")
            if exact_day_sent == "none":
                df.loc[i+1,"exact_day_sent"] = "nday"
            else:
                df.loc[i+1,"exact_day_sent"] = "both"

if now.year - df.loc[1,"this_year"] > 0:
    df.loc[1,"this_year"] = now.year
    for i in range(df.index.max()):
        df.loc[i+1, "week_prior_sent"] = "none"
        df.loc[i+1, "exact_day_sent"] = "none"
df.to_excel(r"C:\Users\Drako\Desktop\Coding\Projects\Birthday reminder\birthday.xlsx")