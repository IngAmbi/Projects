import smtplib
from email.message import EmailMessage

s = smtplib.SMTP('smtp.gmail.com', 587)

s.starttls() #Security TLS
s.login("beranek369@gmail.com", "fywq zlgc chyd mpdp")

# TODO: use f"" to form the message with correct variables from excel doc

msg= EmailMessage()
msg["Subject"] = "Automated email from my Python code"
msg["From"] = "beranek369@gmail.com"
msg["To"] = "beranek.stepan@seznam.cz"
msg.set_content("This is an automated message.")

s.send_message(msg)
s.quit()