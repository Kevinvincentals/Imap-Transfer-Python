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

# Source account details
source_host = input("Enter the source host (IMAP server): ")
source_username = input("Enter the source username: ")
source_password = getpass("Enter the source password: ")

# Destination account details
dest_host = input("Enter the destination host (IMAP server): ")
dest_username = input("Enter the destination username: ")
dest_password = getpass("Enter the destination password: ")

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
            response = source_client.fetch(messages, ['BODY.PEEK[]', 'FLAGS'])

            transferred_count = 0
            duplicate_count = 0

            for msgid in tqdm(response.keys(), desc="Copying emails", unit="email"):
                raw_message = response[msgid][b'BODY[]']
                message = message_from_bytes(raw_message)
                flags = response[msgid][b'FLAGS']

                # Check if the message is already present in the destination mailbox
                message_id = message['Message-ID'].strip() if message['Message-ID'] else ''
                dest_search = dest_client.search(['HEADER', 'Message-ID', message_id]) if message_id else []

                if dest_search:
                    duplicate_count += 1
                else:
                    dest_client.append(dest_box, message.as_bytes(), flags=flags)
                    transferred_count += 1

            print(f"{transferred_count} messages copied from {src_box} to {dest_box}.")
            print(f"{duplicate_count} duplicate messages skipped.")

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
        response = source_client.fetch(messages, ['BODY.PEEK[]', 'FLAGS'])

        transferred_count = 0
        duplicate_count = 0

        for msgid in tqdm(response.keys(), desc="Copying emails", unit="email"):
            raw_message = response[msgid][b'BODY[]']
            message = message_from_bytes(raw_message)
            flags = response[msgid][b'FLAGS']

            # Check if the message is already present in the destination mailbox
            message_id = message['Message-ID'].strip() if message['Message-ID'] else ''
            dest_search = dest_client.search(['HEADER', 'Message-ID', message_id]) if message_id else []
            if dest_search:
                duplicate_count += 1
            else:
                dest_client.append(dest_mailbox, message.as_bytes(), flags=flags)
                transferred_count += 1

        print(f"{transferred_count} messages copied from {source_mailbox} to {dest_mailbox}.")
        print(f"{duplicate_count} duplicate messages skipped.")
