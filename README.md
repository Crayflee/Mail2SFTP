# Mail2SFTP

## Overview

**Mail2SFTP** is a Python-based application that monitors an email inbox for specific messages containing `.csv` file attachments. Upon receiving such emails, it automatically downloads the attachments and securely uploads them to a specified SFTP server. The app is built with a user-friendly GUI using `Tkinter` to provide real-time feedback on the progress and results of the process.

The application also sorts processed emails by moving them to designated folders within the mailbox, ensuring that successful and erroneous operations are correctly handled.

## Features

- **Automated Email Monitoring:** Connects to an IMAP email server and scans for new emails with `.csv` attachments.
- **File Management:** Downloads the `.csv` attachments and securely uploads them to an SFTP server.
- **Error Handling:** Emails without `.csv` attachments are moved to an "Error" folder.
- **Processed Handling:** Emails with successfully processed attachments are moved to a "Processed" folder.
- **GUI Interface:** Provides real-time status updates through a graphical user interface, including progress tracking and error reporting.
- **Execution Logging:** Tracks the last execution date and time, displaying it on the GUI.

## Prerequisites

Before running the application, ensure that you have the following installed:

1. **Python 3.6+**
2. **Required Libraries:**

   Install the required libraries using `pip`:

   ```bash
   pip install -r requirements.txt
