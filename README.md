# Imap-Transfer-Python

This script moves IMAP inboxes between two mail servers. It lets the user specify the source inboxes that need to be moved to specific destination inboxes, giving the user full control. The script **does not** remove the emails from the old account; it simply copies them.

It also uses `tqdm` to display a progress bar with stats and remaining time.

This script has been tested with 60,000 emails, which took around 30 minutes to move.

> ⚠️ **Please note that this script does not create a direct connection from the IMAP server to the IMAP server. It uses the host of the script as a middleman service.**

## Roadmap

~~1. Adding logic to handle if disconnects happen (avoid duplicating mails, etc).~~ (Done)

2. Implement Flask frontend and make the service public

