
# Imap-Transfer-Python (IMAP TOOLS)

Imap-Transfer-Python is a script that enables you to move IMAP inboxes between two mail servers easily. 📨📥

## Transfer Features
* Moves emails from source inboxes to specific destination inboxes, providing full control to the user. 
* Option to use "auto-mode" for automatic identification of source and destination folders. 🔄
* Ensures that duplicate emails are not transferred, preventing duplication if a disconnection occurs during the transfer. 
* Retains email flags
* Supports both SSL and TLS connections, ensuring secure email transfers. 🔒

## Backup Features
* Allows backing up emails locally to password encrypted ZIP file
* Creates a backup of the emails by fetching all message IDs from the source mailbox and storing them in the backup file.

## Restore Features
* Allows restoring emails from a backup file.
* Ensures that duplicate emails are not restored, preventing duplication if a disconnection occurs during the restore process.

This script has been used and tested with over 100 different mailboxes varying from size 200mb to 100GB. The speed of transfer depends on your connection and the mailservices connection. 

Disclaimer
----
> ⚠️ **I do not take any responsbility for the usage of this script**


> ⚠️ **Please note that this script does not create a direct connection from the IMAP server to the IMAP server. It uses the host of the script as a middleman service.**
## Usage/Examples



Follow these steps to get started with Imap-Transfer-Python:

* Make sure you have Python installed on your system.
* Clone the repository to your local machine. (Or server)
* Install the required dependencies by running the following command:

```
pip install -r requirements.txt
```

* Run the transfer.py app, and follow the interactive guide.
## Roadmap

Here are the planned future improvements for Imap-Transfer-Python:


- Implement a Flask web frontend to provide a user-friendly interface.

- Create a public service to allow users to transfer their emails online to a new provider.


## FAQ

#### Will the script remove mails from the source?
No, the script will not remove any mails from the source. It will only take a copy.


## Feedback

If you have any feedback, please reach out to me at kevin@tideo.dk


![Logo](https://tideo.dk/images/logo-light.png)

