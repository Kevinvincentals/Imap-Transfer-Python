# Imap-Transfer-Python

This script moves IMAP inboxes between two mail servers. It lets the user specify the source inboxes that need to be moved to specific destination inboxes, giving the user full control. Furthermore, it also allows to use "auto-mode" where it will try to identify the source/dest folders. The script has logic to not move duplicated emails, so if for some reason a disconnection would happend mid-transfer, it would not transfer already transfered mails.

The script checks for either SSL or TLS connection. You should not move mails over non encrypted connections. 

It uses `tqdm` to display a progress bar with stats and remaining time. You have to install this with `pip install tqdm`

This script has been tested with 60,000 emails, which took around 30 minutes to move.

> ⚠️ **Please note that this script does not create a direct connection from the IMAP server to the IMAP server. It uses the host of the script as a middleman service.**

## Roadmap

~~1. Adding logic to handle if disconnects happen (avoid duplicating mails, etc).~~ (Done)

2. Implement Flask frontend and make the service public



## Roadmap
Usage:
1. Make sure you have python installed

2. Clone the repo

3. Run the transfer.py with python, and the rest will be interactive.