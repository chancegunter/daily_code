import re
import win32com.client
import mysql.connector.connection
import os
import os.path
import sys
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import datetime
from dateutil import relativedelta
from pathlib import Path
cur_dir = Path(__file__).parent
sys.path.append(str(cur_dir / '../../../../../../Credentials'))
import credentials


# ------------------------------------------------------------------------------
# Connects To The DEV DB
# ------------------------------------------------------------------------------


def connect_read():
    conn = mysql.connector.connect(host=credentials.host_read,
                                   database=credentials.database_hq,
                                   user=credentials.username,
                                   password=credentials.password,
                                   ssl_disabled=True)

    return conn

# ===============================================================================
# Gathering all messages
# ===============================================================================


def email_transfer(outlook, mapi, original_path):
    os.chdir(original_path)

    records = []
    directory = r'C:\Project Files\Function_Projects\Custom_Brokerage_Automation\Output'

    email = 'temp.email@host.com'
    category_label = 'exp'
    init_folder = 'init'
    to_folder = 'to'
    messages = mapi.Folders(f"{email}").Folders("fttp " f'{category_label}').Folders(f"{init_folder}").Items
    donebox = mapi.Folders(f"{email}").Folders("fttp " f'{category_label}').Folders(f"{to_folder}")
    folder = 'local_location'

    messages.Sort("[ReceivedTime]", Descending=True)
    mssg_list = list(messages)

    for message in mssg_list:
        attachments = message.Attachments
        attachment_name = message.Subject
        records.append({'Time': message.ReceivedTime, 'Name': message.Subject})

        for attachment in attachments:
            print(folder + '/' + attachment_name)
            attachment.SaveASFile(folder + '/' + attachment_name)
            message.MarkAsTask(5)
            message.Move(donebox)

    final_df = pd.DataFrame.from_records(records)
    final_df['Time'] = final_df['Time'].dt.tz_convert(None)
    final_df.to_csv(directory + '\\' + 'Final_Download.csv', index=False)

    return final_df

# ===============================================================================
# Main
# ===============================================================================


def send_email(original_path, today, final_df):
    os.chdir(original_path)
    os.chdir('../../../../../../Project Files/Function_Projects/Custom_Brokerage_Automation/Output')
    file_name = 'Final_Download.csv'
    ccaddrs = ['ccaddr1@host.com', 'ccaddr2@host.com']

    if len(final_df) > 0:
        fromaddr = 'fromaddr@host.com'
        toaddr = 'toaddr@host.com'
        ccaddr = ', '.join(ccaddrs)
        password = credentials.email_password
        subject = 'subject text' + str(today.strftime('%Y-%m-%d'))
        msg = MIMEMultipart()
        msg['From'] = fromaddr
        msg['To'] = toaddr
        msg['CC'] = ccaddr
        msg['Subject'] = subject
        folder = os.path.abspath(os.curdir)
        spart = MIMEBase('application', "octet-stream")
        spart.set_payload(open(folder + '\\' + file_name, "rb").read())
        encoders.encode_base64(spart)
        spart.add_header('Content-Disposition', f'attachment; filename={file_name}')
        msg.attach(spart)

        body = 'body text'
        msg.attach(MIMEText(body, 'plain'))
        text = msg.as_string()
        server = smtplib.SMTP('smtp-mail.outlook.com', 587)
        server.starttls()
        server.login(fromaddr, password)
        server.sendmail(fromaddr, toaddr, text)
        server.quit()
        print('Email Sent')

    print('Failed_df contains nothing')

# ===============================================================================
# Main
# ===============================================================================


def main():
    original_path = os.getcwd()
    today = datetime.datetime.today()
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace('MAPI')

    final_df = email_transfer(outlook, mapi, original_path)

    send_email(original_path, today, final_df)

# ===============================================================================
# Initialize
# ===============================================================================


if __name__ == '__main__':
    main()
