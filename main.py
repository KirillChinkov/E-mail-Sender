import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText


book=openpyxl.open("List.xlsx", read_only=True)


sheet=book.active

# name=(sheet[2][1].value)
# unit=(sheet[2][0].value)




image = open('unnamed.jpg', 'rb')


def send_email(message):
    sender = "____________@_______________.com"

    password = ""

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()

    try:
        server.login(sender, password)
        msg = MIMEMultipart()
        msg['From']=sender
        msg['To'] = recipient
        msg["Subject"] = "Сообщение о встрече!"

        msg1 = MIMEText(message)
        part = MIMEApplication(open('unnamed.jpg', 'rb').read()) #rb
        msg.attach(part)

        part.add_header('Content-Disposition', 'attachment', filename='unnamed.jpg')
        msg.attach(part)
        msg.attach(msg1)
        server.sendmail(sender, recipient, msg.as_string())

        server.sendmail(sender, recipient, f"Subject: CLICK ME PLEASE!\n{message}")

        return print(f"Message for {recipient} was delivered")
    except Exception as _ex:
        return f"{_ex}\nCheck your login or password please!"





for i in range(2, 3):
    name = (sheet[i][1].value)
    unit = (sheet[i][0].value)
    recipient = (sheet[i][2].value)
    message = f"Здравствуйте, уважаемый(ая) {name}! Мы рады, что Вы приобрели у нас квартиру {unit}!"
    print(message)

    send_email(message)

# def main():
#     message = input("Type your message: ")
#     print(send_email(message=message))
#
#
# if __name__ == "__main__":
#     main()