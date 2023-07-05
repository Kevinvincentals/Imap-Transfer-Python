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

    source_mailbox = choose_mailbox(source_client, "Source mailboxes:")
    dest_mailbox = choose_mailbox(dest_client, "Destination mailboxes:")

    # Select the mailbox you want to copy from in the source account
    select_info = source_client.select_folder(source_mailbox)

    # Fetch all message id's from source mailbox
    messages = source_client.search('ALL')

    print(f"\nTotal emails to be copied: {len(messages)}")

    # Get the full message data (including FLAGS) for the messages
    response = source_client.fetch(messages, ['BODY.PEEK[]', 'FLAGS'])

    for msgid in tqdm(response.keys(), desc="Copying emails", unit="email"):
        raw_message = response[msgid][b'BODY[]']
        message = message_from_bytes(raw_message)
        flags = response[msgid][b'FLAGS']

        # Save the messages to the destination mailbox with same flags
        dest_client.append(dest_mailbox, message.as_bytes(), flags=flags)

    print(f"\n{len(response)} messages copied from {source_mailbox} to {dest_mailbox}.")
