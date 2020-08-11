import argparse
import smtplib
import ssl
import os
from email.message import EmailMessage
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders


def get_args():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '--send-method',
        dest='send_method',
        choices={
            'ssl',
            'tls'},
        default='tls',
        required=False)
    parser.add_argument(
        '--smtp-host',
        dest='smtp_host',
        required=True)
    parser.add_argument(
        '--smtp-port',
        dest='smtp_port',
        default='',
        required=True)
    parser.add_argument(
        '--sender-address',
        dest='sender_address',
        default='',
        required=True)
    parser.add_argument(
        '--sender-name',
        dest='sender_name',
        default='',
        required=False)
    parser.add_argument(
        '--to',
        dest='to',
        default='',
        required=False)
    parser.add_argument(
        '--cc',
        dest='cc',
        default='',
        required=False)
    parser.add_argument(
        '--bcc',
        dest='bcc',
        default='',
        required=False)
    parser.add_argument(
        '--password',
        dest='password',
        default='',
        required=True)
    parser.add_argument(
        '--subject',
        dest='subject',
        default='',
        required=False)
    parser.add_argument(
        '--message',
        dest='message',
        default='',
        required=True)
    parser.add_argument(
        '--source-file-name',
        dest='source_file_name',
        default='',
        required=False)
    parser.add_argument(
        '--source-folder-name',
        dest='source_folder_name',
        default='',
        required=False)
    return parser.parse_args()


def create_message_object(
        sender_address,
        sender_name,
        to,
        cc,
        bcc,
        subject,
        message):
    """
    Create an Message object, msg, by using the provided send parameters.
    """
    msg = MIMEMultipart()

    msg['Subject'] = subject
    msg['From'] = f'{sender_name}<{sender_address}>'
    msg['To'] = to
    msg['Cc'] = cc
    msg['Bcc'] = bcc

    msg.attach(MIMEText(message, "html"))

    return msg


def add_attachment_to_message_object(msg, source_file_path):
    """
    Add source_file_path as an attachment to the message object.
    """
    try:
        complete_record = MIMEBase('application', 'octet-stream')
        complete_record.set_payload((open(source_file_path, "rb").read()))
        encoders.encode_base64(complete_record)
        complete_record.add_header(
            'Content-Disposition',
            'attachment',
            filename=os.path.basename(source_file_path))
        msg.attach(complete_record)

        upload_record = MIMEBase('application', 'octet-stream')
        upload_record.set_payload((open(source_file_path, "rb").read()))
        encoders.encode_base64(upload_record)
        upload_record.add_header(
            'Content-Disposition',
            'attachment',
            filename=os.path.basename(source_file_path))
        msg.attach(upload_record)
        return msg
    except Exception as e:
        raise(e)
        print("Could not attach the files to the email.")


def send_tls_message(
        smtp_host,
        smtp_port,
        sender_address,
        password,
        msg):
    """
    Send an email using the TLS connection method.
    """
    context = ssl.create_default_context()
    try:
        server = smtplib.SMTP(smtp_host, smtp_port)
        server.starttls(context=context)
        server.login(sender_address, password)
        server.send_message(msg)
        server.quit()
    except Exception as e:
        raise(e)


def send_ssl_message(
        smtp_host,
        smtp_port,
        sender_address,
        password,
        msg):
    """
    Send an email using the SSL connection method.
    """
    context = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL(smtp_host, smtp_port, context=context) as server:
            server.login(sender_address, password)
            server.send_message(msg)
    except Exception as e:
        raise(e)


def clean_folder_name(folder_name):
    """
    Cleans folders name by removing duplicate '/' as well as leading and trailing '/' characters.
    """
    folder_name = folder_name.strip('/')
    if folder_name != '':
        folder_name = os.path.normpath(folder_name)
    return folder_name


def combine_folder_and_file_name(folder_name, file_name):
    """
    Combine together the provided folder_name and file_name into one path variable.
    """
    combined_name = os.path.normpath(
        f'{folder_name}{"/" if folder_name else ""}{file_name}')
    combined_name = os.path.normpath(combined_name)

    return combined_name


def create_shipyard_link():
    """
    Create a link back to the Shipyard log page for the current alert.
    """
    org_name = os.environ.get('SHIPYARD_ORG_NAME')
    project_id = os.environ.get('SHIPYARD_PROJECT_ID')
    vessel_id = os.environ.get('SHIPYARD_VESSEL_ID')
    log_id = os.environ.get('SHIPYARD_LOG_ID')

    if org_name and project_id and vessel_id and log_id:
        shipyard_link = f'https://app.shipyardapp.com/{org_name}/projects/{project_id}/vessels/{vessel_id}/logs/{log_id}'
    else:
        shipyard_link = 'https://www.shipyardapp.com'
    return shipyard_link


def add_shipyard_link_to_message(message, shipyard_link):
    """
    Create a "signature" at the bottom of the email that links back to Shipyard.
    """
    message = f'{message}<br><br>---<br>Sent by <a href=https://www.shipyardapp.com> Shipyard</a> | <a href={shipyard_link}>Click Here</a> to Edit'
    return message


def main():
    args = get_args()
    send_method = args.send_method
    smtp_host = args.smtp_host
    smtp_port = int(args.smtp_port)
    sender_address = args.sender_address
    sender_name = args. sender_name
    to = args.to
    cc = args.cc
    bcc = args.bcc
    password = args.password
    subject = args.subject
    message = args.message

    source_file_name = args.source_file_name
    source_folder_name = clean_folder_name(args.source_folder_name)
    source_full_path = combine_folder_and_file_name(
        folder_name=source_folder_name, file_name=source_file_name)

    shipyard_link = create_shipyard_link()
    message = add_shipyard_link_to_message(message, shipyard_link)

    msg = create_message_object(
        sender_address,
        sender_name,
        to,
        cc,
        bcc,
        subject,
        message)

    if source_file_name:
        msg = add_attachment_to_message_object(msg, source_full_path)

    if send_method == 'ssl':
        send_ssl_message(
            smtp_host,
            smtp_port,
            sender_address,
            password,
            msg)
    else:
        send_tls_message(
            smtp_host,
            smtp_port,
            sender_address,
            password,
            msg)


if __name__ == '__main__':
    main()
