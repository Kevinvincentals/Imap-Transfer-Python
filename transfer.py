import os
import zipfile
from getpass import getpass
from imapclient import IMAPClient
from email import message_from_bytes
from tqdm import tqdm

def choose_mailbox(client, prompt):
    """Let the user choose a mailbox from the given IMAPClient instance."""
    mailboxes = [box[-1] for box in client.list_folders()]
    while True:
        print(f"\n{prompt}")
        for i, mailbox in enumerate(mailboxes, 1):
            print(f"{i}. {mailbox}")
        choice = input("Choose a mailbox (by number): ")
        if choice.isdigit() and 1 <= int(choice) <= len(mailboxes):
            return mailboxes[int(choice) - 1]
        print("Invalid choice, please try again.")

def match_mailboxes(source_client, dest_client):
    """Match source and destination mailboxes based on their names."""
    source_mailboxes = [box[-1] for box in source_client.list_folders()]
    dest_mailboxes = [box[-1] for box in dest_client.list_folders()]

    matches = []
    for src_box in source_mailboxes:
        for dest_box in dest_mailboxes:
            if src_box.lower() == dest_box.lower():
                matches.append((src_box, dest_box))
                break
    return matches

# Backup option
backup_option = input("Choose an option:\n1. Transfer emails\n2. Backup emails\nEnter your choice: ")
while backup_option not in ['1', '2', '3']:
    print("Invalid choice, please try again.")
    backup_option = input("Choose an option:\n1. Transfer emails\n2. Backup emails\nEnter your choice: ")


# Source account details
source_host = input("Enter the source host (IMAP server): ")
source_username = input("Enter the source username: ")
source_password = getpass("Enter the source password: ")

# Destination account details if transfer option is chosen
if backup_option == '1':
    dest_host = input("Enter the destination host (IMAP server): ")
    dest_username = input("Enter the destination username: ")
    dest_password = getpass("Enter the destination password: ")

# Prepare backup file if backup option is chosen
if backup_option == '2':
    backup_password = getpass("Enter a password for the backup file: ")
    backup_filename = "email_backup.zip"
    backup_mailboxes = []

# Function to try to connect with SSL, then with TLS if SSL fails
def connect_imap(host, username, password):
    for port, ssl in [(993, True), (143, False)]:
        try:
            client = IMAPClient(host, port=port, use_uid=True, ssl=ssl)
            client.login(username, password)
            return client
        except:
            continue
    raise ValueError(f"Could not connect to {host} with the provided credentials")

if backup_option == '1':
    # Connect to the servers
    with connect_imap(source_host, source_username, source_password) as source_client, \
         connect_imap(dest_host, dest_username, dest_password) as dest_client:

        auto_move = input("Do you want to automatically match mailboxes and move emails? (y/n): ").lower().strip() == "y"

        if auto_move:
            matches = match_mailboxes(source_client, dest_client)
            if not matches:
                print("No matching mailboxes found.")
                exit()

            for src_box, dest_box in matches:
                print(f"\nMoving emails from {src_box} to {dest_box}...")
                source_client.select_folder(src_box)
                dest_client.select_folder(dest_box)

                # Fetch all message IDs from source mailbox
                source_messages = source_client.search('ALL')

                # Fetch all message IDs from destination mailbox
                dest_messages = dest_client.search('ALL')

                # Get the message IDs already present in the destination mailbox
                dest_message_ids = set(dest_messages)

                # Filter out the already transferred messages from the source mailbox
                messages = [msg_id for msg_id in source_messages if msg_id not in dest_message_ids]

                print(f"Total emails to be copied: {len(messages)}")

                if len(messages) > 4000:
                    print("Due to the large quantity of emails, this may take some time. Please wait...")

                # Get the full message data (including FLAGS) for the messages
                response = source_client.fetch(messages, ['BODY.PEEK[]', 'FLAGS', 'RFC822.SIZE'])

                transferred_count = 0
                duplicate_count = 0
                total_size = 0

                with tqdm(total=len(response), desc="Copying emails", unit="email", ncols=80) as pbar:
                    for msgid in response.keys():
                        raw_message = response[msgid][b'BODY[]']
                        message = message_from_bytes(raw_message)
                        flags = response[msgid][b'FLAGS']
                        size = response[msgid][b'RFC822.SIZE']
                        total_size += size

                        # Check if the message is already present in the destination mailbox
                        message_id = message['Message-ID'].strip() if message['Message-ID'] else ''
                        dest_search = dest_client.search(['HEADER', 'Message-ID', message_id]) if message_id else []

                        if dest_search:
                            duplicate_count += 1
                        else:
                            dest_client.append(dest_box, message.as_bytes(), flags=flags)
                            transferred_count += 1

                        pbar.set_postfix({"Size moved:": f"{total_size/(1024*1024):.2f} MB"})
                        pbar.update(1)

                print(f"\n{transferred_count} messages copied from {src_box} to {dest_box}.")
                print(f"{duplicate_count} duplicate messages skipped.")
                print(f"Total size of moved emails: {total_size / (1024 * 1024):.2f} MB")

        else:
            source_mailbox = choose_mailbox(source_client, "Source mailboxes:")
            dest_mailbox = choose_mailbox(dest_client, "Destination mailboxes:")

            source_client.select_folder(source_mailbox)
            dest_client.select_folder(dest_mailbox)

            # Fetch all message IDs from source mailbox
            source_messages = source_client.search('ALL')

            # Fetch all message IDs from destination mailbox
            dest_messages = dest_client.search('ALL')

            # Get the message IDs already present in the destination mailbox
            dest_message_ids = set(dest_messages)

            # Filter out the already transferred messages from the source mailbox
            messages = [msg_id for msg_id in source_messages if msg_id not in dest_message_ids]

            print(f"\nTotal emails to be copied: {len(messages)}")

            if len(messages) > 4000:
                print("Due to the large quantity of emails, this may take some time. Please wait...")

            # Get the full message data (including FLAGS) for the messages
            response = source_client.fetch(messages, ['BODY.PEEK[]', 'FLAGS', 'RFC822.SIZE'])

            transferred_count = 0
            duplicate_count = 0
            total_size = 0

            with tqdm(total=len(response), desc="Copying emails", unit="email", ncols=80) as pbar:
                    for msgid in response.keys():
                        raw_message = response[msgid][b'BODY[]']
                        message = message_from_bytes(raw_message)
                        flags = response[msgid][b'FLAGS']
                        size = response[msgid][b'RFC822.SIZE']
                        total_size += size

                        # Check if the message is already present in the destination mailbox
                        message_id = message['Message-ID'].strip() if message['Message-ID'] else ''
                        dest_search = dest_client.search(['HEADER', 'Message-ID', message_id]) if message_id else []

                        if dest_search:
                            duplicate_count += 1
                        else:
                            dest_client.append(dest_mailbox, message.as_bytes(), flags=flags)
                            transferred_count += 1

                        pbar.set_postfix({"Transferred": transferred_count, "Duplicates": duplicate_count, "Total Size": f"{total_size/(1024*1024):.2f} MB"})
                        pbar.update(1)

            print(f"\n{transferred_count} messages copied from {source_mailbox} to {dest_mailbox}.")
            print(f"{duplicate_count} duplicate messages skipped.")
            print(f"Total size of moved emails: {total_size / (1024 * 1024):.2f} MB")

        # Summary
        print("\n--- Summary ---")
        if auto_move:
            for src_box, dest_box in matches:
                print(f"Moved from Source {src_box} to Destination {dest_box}:")
                print(f"  - Emails moved: {transferred_count}")
                print(f"  - Duplicate messages skipped: {duplicate_count}")
                print(f"  - Total size: {total_size / (1024 * 1024):.2f} MB")
                print()
        else:
            print(f"Moved from Source {source_mailbox} to Destination {dest_mailbox}:")
            print(f"  - Emails moved: {transferred_count}")
            print(f"  - Duplicate messages skipped: {duplicate_count}")
            print(f"  - Total size: {total_size / (1024 * 1024):.2f} MB")

elif backup_option == '2':
    # Connect to the source server for backup
    with connect_imap(source_host, source_username, source_password) as source_client:
        # Get the list of source mailboxes
        source_mailboxes = [box[-1] for box in source_client.list_folders()]

        # Select specific mailboxes for backup
        print("\nSelect the mailboxes to backup:")
        for i, mailbox in enumerate(source_mailboxes, 1):
            print(f"{i}. {mailbox}")
        print("Enter the mailbox numbers to backup (comma-separated): ")
        choices = input("Choose mailboxes: ").split(",")
        for choice in choices:
            choice = choice.strip()
            if choice.isdigit() and 1 <= int(choice) <= len(source_mailboxes):
                backup_mailboxes.append(source_mailboxes[int(choice) - 1])
            else:
                print(f"Invalid choice: {choice}. Skipping.")

        if not backup_mailboxes:
            print("No mailboxes selected for backup. Exiting.")
            exit()

        # Connect to the source server again to perform backup
        with connect_imap(source_host, source_username, source_password) as source_client:
            print(f"\nBacking up {len(backup_mailboxes)} mailboxes...")

            with zipfile.ZipFile(backup_filename, "w", zipfile.ZIP_DEFLATED) as backup_zip:
                backup_count = 0

                for mailbox in tqdm(backup_mailboxes, desc="Backing up mailboxes", ncols=80):
                    source_client.select_folder(mailbox)

                    # Fetch all message IDs from the mailbox
                    messages = source_client.search('ALL')

                    # Fetch the full message data for the messages
                    response = source_client.fetch(messages, ['BODY.PEEK[]'])

                    for msgid, data in response.items():
                        raw_message = data[b'BODY[]']
                        message = message_from_bytes(raw_message)

                        # Create a folder for the current message within the backup zip file
                        message_folder = f"{mailbox}/{msgid}"
                        backup_zip.writestr(f"{message_folder}/message.eml", raw_message)

                        # Save attachments
                        for part in message.walk():
                            if part.get_content_maintype() == "multipart":
                                continue
                            if part.get("Content-Disposition") is None:
                                continue

                            filename = part.get_filename()
                            if filename:
                                attachment_data = part.get_payload(decode=True)
                                backup_zip.writestr(f"{message_folder}/{filename}", attachment_data)

                        backup_count += 1

            print(f"\nBackup created successfully: {backup_filename}")
            print(f"Total emails backed up: {backup_count}")
            print('Please note, there is currently no logic to handle importing of emails')
