import argparse
import smtplib
import ssl
import os
import re
from email.message import EmailMessage
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
import shipyard_utils as shipyard
from tabulate import tabulate


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
        '--username',
        dest='username',
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
    parser.add_argument(
        '--include-shipyard-footer',
        dest='include_shipyard_footer',
        default='TRUE',
        choices={
            'TRUE',
            'FALSE'},
        required=False)

    args = parser.parse_args()
    if not (args.to or args.cc or args.bcc):
        parser.error(
            'Email requires at least one recepient using --to, --cc, or --bcc')
    return args

def _has_file(message:str) -> bool:
    """ Returns true if a message string has the {{file.txt}} pattern

    Args:
        message (str): The message

    Returns:
        bool: 
    """
    pattern = r'\{\{[^\{\}]+\}\}'
    res = re.search(pattern,message)
    if res is not None:
        return True
    return False

def _extract_file(message:str) -> str:
    pattern = r'\{\{[^\{\}]+\}\}'
    res = re.search(pattern,message).group()
    file_pattern = re.compile(r'[{}]+')
    return re.sub(file_pattern, '', res)

def _read_file(file:str, message:str) -> str:
    try:
        with open(file, 'r') as f:
            content = f.read()
            f.close()
    except Exception as e:
        print(f"Could not load the contents of file {file}. Make sure the file extension is provided")
        raise(FileNotFoundError)
    pattern = r'\{\{[^\{\}]+\}\}'
    msg = re.sub('\n','<br>',f"{re.sub(pattern,'',message)} <br><br> {content}")
    return msg

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
        print(e)
        print("Could not attach the files to the email.")


def send_tls_message(
        smtp_host,
        smtp_port,
        username,
        password,
        msg):
    """
    Send an email using the TLS connection method.
    """
    context = ssl.create_default_context()
    try:
        server = smtplib.SMTP(smtp_host, smtp_port)
        server.starttls(context=context)
        server.login(username, password)
        server.send_message(msg)
        server.quit()
        print('Message successfully sent.')
    except Exception as e:
        raise(e)


def send_ssl_message(
        smtp_host,
        smtp_port,
        username,
        password,
        msg):
    """
    Send an email using the SSL connection method.
    """
    context = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL(smtp_host, smtp_port, context=context) as server:
            server.login(username, password)
            server.send_message(msg)
        print('Message successfully sent.')
    except Exception as e:
        raise(e)


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
        file_names = shipyard.files.find_all_local_file_names(
            source_folder_name)
        matching_file_names = shipyard.files.find_all_file_matches(
            file_names, re.compile(source_file_name))

        files_to_upload = matching_file_names
    else:
        source_full_path = shipyard.files.combine_folder_and_file_name(
            folder_name=source_folder_name, file_name=source_file_name)
        files_to_upload = [source_full_path]
    return files_to_upload


def should_message_be_sent(
        conditional_send,
        source_full_path,
        source_file_name_match_type):
    """
    Determine if an email message should be sent based on the parameters provided.
    """

    if source_file_name_match_type == 'exact_match':
        if (
                conditional_send == 'file_exists' and os.path.exists(
                    source_full_path[0])) or (
                conditional_send == 'file_dne' and not os.path.exists(
                    source_full_path[0])) or (
                conditional_send == 'always'):
            return True
        else:
            return False
    elif source_file_name_match_type == 'regex_match':
        if (
            conditional_send == 'file_exists' and len(source_full_path) > 0) or (
                conditional_send == 'file_dne' and len(source_full_path) == 0) or (
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
    username = args.username
    password = args.password
    subject = args.subject
    message = args.message
    conditional_send = args.conditional_send
    source_file_name_match_type = args.source_file_name_match_type
    file_upload = args.file_upload
    include_shipyard_footer = shipyard.args.convert_to_boolean(
        args.include_shipyard_footer)

    if not username:
        username = sender_address

    source_file_name = args.source_file_name
    source_folder_name = shipyard.files.clean_folder_name(
        args.source_folder_name)

    file_paths = determine_file_to_upload(
        source_file_name_match_type=source_file_name_match_type,
        source_folder_name=source_folder_name,
        source_file_name=source_file_name)

    if should_message_be_sent(
            conditional_send,
            file_paths,
            source_file_name_match_type):

        if _has_file(message):
            file = _extract_file(message)
            message = _read_file(file, message)
        
        if include_shipyard_footer:
            shipyard_link = shipyard.args.create_shipyard_link()
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

            if shipyard.files.are_files_too_large(
                    file_paths, max_size_bytes=10000000):
                compressed_file_name = shipyard.files.compress_files(
                    file_paths,
                    destination_full_path='Archive.zip',
                    compression='zip')
                print(f'Attaching {compressed_file_name} to message.')
                msg = add_attachment_to_message_object(
                    msg=msg, file_path=compressed_file_name)
            else:
                for file in file_paths:
                    print(f'Attaching {file} to message.')
                    msg = add_attachment_to_message_object(
                        msg=msg, file_path=file)

        if send_method == 'ssl':
            send_ssl_message(
                smtp_host=smtp_host,
                smtp_port=smtp_port,
                username=username,
                password=password,
                msg=msg)
        else:
            send_tls_message(
                smtp_host,
                smtp_port=smtp_port,
                username=username,
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
