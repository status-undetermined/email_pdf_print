import imaplib, ssl, email, traceback
import xml.etree.ElementTree as ET
from subprocess import Popen, PIPE

# Python 3.12
# OS: Linux
# cups and lp required



# Set XML-config-file
tree = ET.parse('config.xml')
root = tree.getroot()

# Get required variables from given XML-file. All variables MUST be set, except 'coding'
for entry in root.findall('./account'):
    acc_name = entry.attrib['name']
    mail_user = entry.find('mail/mail_user').text
    mail_pass = entry.find('mail/mail_pass').text
    mail_host = entry.find('mail/mail_host').text
    mail_port = int(entry.find('mail/mail_port').text)
    mail_search_criteria = entry.find('mail/mail_search_criteria').text
    mail_mailbox = entry.find('mail/mail_mailbox').text
    mail_processed = entry.find('mail/mail_processed').text
    coding = entry.find('attachment/coding').text
    printer_name = entry.find('printer/printer_name').text
    print(mail_search_criteria)



########################################################################################################################



    ### new instance of IMAP4_SSL "mail", login at mailserver
    try:
        mail = imaplib.IMAP4_SSL(mail_host, mail_port, ssl_context=ssl.SSLContext(ssl.PROTOCOL_SSLv23))
    except Exception:
        print(traceback.format_exc())
        print('Config or Server Error in Account: ', end='')
        print(acc_name)
        continue


    try:
        mail.login(mail_user, mail_pass)
    except Exception:
        print(traceback.format_exc())
        print('Config or Server Error in Account: ', end='')
        print(acc_name)
        continue




    ### select mailbox
    status, mail_count = mail.select(mail_mailbox, False)
    if status == 'NO':
        print(traceback.format_exc())
        print('Mailbox error: ', end='')
        print(acc_name, end=' - ')
        print(mail_count)
        continue



    ### get all messages in mailbox with defined search criteria
    typ, data = mail.search(None, mail_search_criteria)
    mail_ids = data[0].decode()
    print("Matching mail_ids")
    print(mail_ids)



    ### iterate over data[0] to get mail ids (VAR mail_ids)
    # fetch mail by id
    # write raw mail content to variable and decode
    # create message object from string

    for mid in data[0].split():
        try:
            typ, data = mail.fetch(mid, '(BODY.PEEK[])')
            raw_mail_content = data[0][1].decode()
            mail_message = email.message_from_string(raw_mail_content)
        except Exception:
            print(traceback.format_exc())
            print('Unable to open message. mail-ID: ', end='')
            print(mid, end='')
            print(' in ', end='')
            print(mail_mailbox, end='')
            print(' from account: ', end='')
            print(acc_name)
            continue


    ### iterate over mail_message and get every message part
        for part in mail_message.walk():
            if part.get_content_subtype() == 'pdf' or (part.get('Content-Disposition') is not None and (part.get('Content-Disposition').split()[0] == 'attachment;' or part.get('Content-Disposition').split()[0] == 'attachment')):
                try:
                    filename = (part.get('Content-Disposition').split("name=")[1][1:]).split('"')[0]
                except Exception:
                    try:
                        filename = (part.get('Content-Type').split("name=")[1][1:]).split('"')[0]
                    except Exception:
                        print(traceback.format_exc())
                        print("Could not fetch filename")
                        continue

                print(filename)
                if filename.endswith('.pdf') is False: continue
                attachment = part.get_payload(decode=True)
                if attachment is None: continue
                # check if attachment is PDF
                if bytes(attachment).startswith(b'%PDF') is not True: continue



                ### Open subprocess object "printer", open stdin and stdout PIPES to subprocess (lp)
                try:
                    printer = Popen(["/usr/bin/lp", '-d', printer_name, '-t', filename, '-'], stdout=PIPE, stdin=PIPE)
                    # write decoded attachment to stdin for subprocess "lp"
                    printer.stdin.write(attachment)
                    # close stdin-PIPE to signal the subprocess to continue and print
                    printer.stdin.close()
                    pr_stdout = printer.stdout.read().decode()
                    printer.stdout.close()
                    print(pr_stdout)
                except Exception:
                    print(traceback.format_exc())
                    print('Unable to print for Account: ', end='')
                    print(acc_name)


                # deconstruct "printer" object
                del printer



                # Mark current mail as seen, copy to mail_processed directory.
                # If mail_mailbox equals mail_processed, mails will not be copied and just flagged as seen.
                mail.store(mid, 'FLAGS', '\\Seen')
                if mail_mailbox != mail_processed:
                    mail.copy(mid, mail_processed)
                # Uncomment following line to set deletet FLAG
                #mail.store(mid, 'FLAGS', '\\Deleted')
    # Uncomment following line to delete FLAGGED messages
    mail.expunge()
    mail.close()
    mail.logout()
    del mail
print("done")