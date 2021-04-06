import argparse
import smtplib
import ssl
import os
import glob
import re
import urllib.parse
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
    parser.add_argument(
        '--conditional-send',
        dest='conditional_send',
        default='always',
        required=False,
        choices={
            'file_exists',
            'file_dne',
            'always'})
    parser.add_argument(
        '--source-file-name-match-type',
        dest='source_file_name_match_type',
        default='exact_match',
        choices={
            'exact_match',
            'regex_match'},
        required=False)
    parser.add_argument(
        '--file-upload',
        dest='file_upload',
        default='no',
        required=True,
        choices={
            'yes',
            'no'})

    args = parser.parse_args()
    if not (args.to or args.cc or args.bcc):
        parser.error(
            'Email requires at least one recepient using --to, --cc, or --bcc')
    return args


def create_message_object(
        sender_address,
        message,
        sender_name=None,
        to=None,
        cc=None,
        bcc=None,
        subject=None):
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


def add_attachment_to_message_object(msg, file_path):
    """
    Add source_file_path as an attachment to the message object.
    """
    try:
        upload_record = MIMEBase('application', 'octet-stream')
        upload_record.set_payload((open(file_path, "rb").read()))
        encoders.encode_base64(upload_record)
        upload_record.add_header(
            'Content-Disposition',
            'attachment',
            filename=os.path.basename(file_path))
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
        print('Message successfully sent.')
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
        print('Message successfully sent.')
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

    if project_id and vessel_id and log_id:
        dynamic_link_section = urllib.parse.quote(
            f'{org_name}/projects/{project_id}/vessels/{vessel_id}/logs/{log_id}')
        shipyard_link = f'https://app.shipyardapp.com/{dynamic_link_section}'
    else:
        shipyard_link = 'https://www.shipyardapp.com'
    return shipyard_link


def add_shipyard_link_to_message(message, shipyard_link):
    """
    Create a "signature" at the bottom of the email that links back to Shipyard.
    """
    message = f'{message}<br><br>---<br>Sent by <a href=https://www.shipyardapp.com> Shipyard</a> | <a href={shipyard_link}>Click Here</a> to Edit'
    return message


def determine_file_to_upload(
        source_file_name_match_type,
        source_folder_name,
        source_file_name):
    """
    Determine whether the file name being uploaded to email
    will be named archive_file_name or will be the source_file_name provided.
    """
    if source_file_name_match_type == 'regex_match':
        file_names = find_all_local_file_names(source_folder_name)
        matching_file_names = find_all_file_matches(
            file_names, re.compile(source_file_name))

        files_to_upload = matching_file_names
    else:
        source_full_path = combine_folder_and_file_name(
            folder_name=source_folder_name, file_name=source_file_name)
        files_to_upload = [source_full_path]
    return files_to_upload


def find_all_local_file_names(source_folder_name):
    """
    Returns a list of all files that exist in the current working directory,
    filtered by source_folder_name if provided.
    """
    cwd = os.getcwd()
    cwd_extension = os.path.normpath(f'{cwd}/{source_folder_name}/**')
    file_names = glob.glob(cwd_extension, recursive=True)
    return file_names


def find_all_file_matches(file_names, file_name_re):
    """
    Return a list of all file_names that matched the regular expression.
    """
    matching_file_names = []
    for file in file_names:
        if re.search(file_name_re, file):
            matching_file_names.append(file)

    return matching_file_names


def should_message_be_sent(
        conditional_send,
        source_folder_name,
        source_file_name,
        source_file_name_match_type):
    """
    Determine if an email message should be sent based on the parameters provided.
    """

    source_full_path = combine_folder_and_file_name(
        source_folder_name, source_file_name)

    if source_file_name_match_type == 'exact_match':
        if (
            conditional_send == 'file_exists' and os.path.exists(source_full_path)) or (
            conditional_send == 'file_dne' and not os.path.exists(source_full_path)) or (
                conditional_send == 'always'):
            return True
        else:
            return False
    if source_file_name_match_type == 'regex_match':
        file_names = find_all_local_file_names(source_folder_name)
        matching_file_names = find_all_file_matches(
            file_names, re.compile(source_file_name))
        if (
            conditional_send == 'file_exists' and len(matching_file_names) > 0) or (
                conditional_send == 'file_dne' and len(matching_file_names) == 0) or (
                    conditional_send == 'always'):
            return True
        else:
            return False


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
    conditional_send = args.conditional_send
    source_file_name_match_type = args.source_file_name_match_type
    file_upload = args.file_upload

    source_file_name = args.source_file_name
    source_folder_name = clean_folder_name(args.source_folder_name)
    source_full_path = combine_folder_and_file_name(
        folder_name=source_folder_name, file_name=source_file_name)

    if should_message_be_sent(
            conditional_send,
            source_folder_name,
            source_file_name,
            source_file_name_match_type):

        shipyard_link = create_shipyard_link()
        message = add_shipyard_link_to_message(
            message=message, shipyard_link=shipyard_link)

        msg = create_message_object(
            sender_address=sender_address,
            message=message,
            sender_name=sender_name,
            to=to,
            cc=cc,
            bcc=bcc,
            subject=subject)

        if file_upload == 'yes':
            files_to_upload = determine_file_to_upload(
                source_file_name_match_type=source_file_name_match_type,
                source_folder_name=source_folder_name,
                source_file_name=source_file_name)
            for file in files_to_upload:
                msg = add_attachment_to_message_object(
                    msg=msg, file_path=file)

        if send_method == 'ssl':
            send_ssl_message(
                smtp_host=smtp_host,
                smtp_port=smtp_port,
                sender_address=sender_address,
                password=password,
                msg=msg)
        else:
            send_tls_message(
                smtp_host,
                smtp_port=smtp_port,
                sender_address=sender_address,
                password=password,
                msg=msg)
    else:
        if conditional_send == 'file_exists':
            print('File(s) could not be found. Message not sent.')
        if conditional_send == 'file_dne':
            print(
                'File(s) were found, but message was conditional based on file not existing. Message not sent.')


if __name__ == '__main__':
    main()
