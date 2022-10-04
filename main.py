import smtplib
from email.message import EmailMessage
import ssl
from datetime import date
import excel_tracker


def mail_automation(person_name, receiver_mail):
    email_sender = "mohammedali.wishes@gmail.com"
    password = 'yfbxtssakiuobxhx'
    email_receiver = receiver_mail
    person_name = person_name

    subject = "Happy Birthday "+person_name+" from Ali"
    body = """
    To a good friend and distinguished colleague, Happy Birthday {0}
    No words can express the respect and admiration I have for you and hope your birthday is a special day.
    I just want to tell you what you are so special to me and how much I value for our relationship. 
    Please remain the person you are today, and time will never put its mark on your beautiful soul.
    
    Enjoy the day and have fun.
    
    From your friend
    Mohammed Ali...
    """.format(person_name)

    email_obj = EmailMessage()
    email_obj['Subject'] = subject
    email_obj['From'] = email_sender
    email_obj['To'] = email_receiver
    email_obj.set_content(body)

    context = ssl.create_default_context()
    port = 465
    with smtplib.SMTP_SSL('smtp.gmail.com', port, context=context) as server:
        server.login(email_sender, password)
        server.sendmail(email_sender, email_receiver, email_obj.as_string())
        server.close()


def today():
    now_day, now_month = date.today().day, date.today().month
    return now_day, now_month


def get_receiver_info(date, month):
    birthday_list = excel_tracker.excel_extract(date, month)
    return birthday_list


if __name__ == '__main__':
    # My email dedicated for birthday wishes
    sender_mail = 'psmohammedalibackupvit@gmail.com'

    # Current date and month
    date, month = today()

    # Retrieve birthday list for the current date and month from excel sheet
    birthday_list = get_receiver_info(date, month)

    for curr_person in birthday_list:
        mail_automation(curr_person['name'], curr_person['email_id'])

# pmshaikabdulla @ gmail.com
# pmshaikabdulla@gmail.com
