import csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

'''Program to mass sending e-mail with attachment ex. invoice for our clients'''
# set path to directory with documents
path_with_files = '/Users/example/example/dir/'
# set path tp CSV file with e-mail address and file name to sending
CSV_path = '/Users/example/example/dir/List.csv'
# set SMTP server
SMTP_server = 'outlook.office365.com'
# set login to SMTP_server
SMTP_login = 'example@example.com'
# set pass to SMTP_server
SMTP_pass = 'example123!'


def send_email(recipient, CSV_list):
    # message option
    msg = MIMEMultipart()
    msg['Subject'] = 'Test subject e-mail'
    msg['From'] = 'example@example.com'
    msg['To'] = recipient
    message = 'Hello, This is test message'
    msg.attach(MIMEText(message))
    filename = CSV_list
    # path to files
    fp = open(path_with_files + filename, 'rb')
    att = MIMEApplication(fp.read(), _subtype="pdf")
    fp.close()
    att.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(att)

    # server smtp settings
    server = smtplib.SMTP(SMTP_server)
    server.starttls()
    server.login(SMTP_login, SMTP_pass)
    server.send_message(msg)
    server.quit()


file = open(CSV_path)
reader = csv.reader(file, delimiter=';')
data = list(reader)

# sending email with exceptions
for client in data:
    try:
        send_email(client[1], client[0])
        print("sanded e-mail to ", client)
    except FileNotFoundError:
        print('No file for: ', client[1], client[0])
    except smtplib.SMTPAuthenticationError:
        print('Wrong login to SMTP server')