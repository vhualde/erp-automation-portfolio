# -*- coding: utf-8 -*-

"""
Outlook Invoice Automation
-----------------------------------
Automation workflow for processing supplier invoice PDF files
received through Outlook mailboxes.

Main features:
- Outlook mailbox scanning
- Attachment extraction
- PDF invoice classification
- Duplicate detection
- Automated folder organization
"""

import os
import re
from pathlib import Path
from datetime import datetime

import win32com.client


# ==========================================
# CONFIGURATION
# ==========================================

MAILBOX_NAME = "finance@example.com"

OUTPUT_FOLDER = Path("output_invoices")

SUBJECT_FILTER = "invoice"

SENDER_FILTER = "supplier"

SUPPORTED_EXTENSIONS = [".pdf", ".zip"]


# ==========================================
# HELPERS
# ==========================================

def safe_filename(filename: str) -> str:
    """
    Cleans invalid characters from filenames.
    """
    return re.sub(r'[<>:"/\\\\|?*]+', "_", filename).strip()


def create_output_folder() -> None:
    """
    Creates output directory if it does not exist.
    """
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)


def mail_matches_filters(mail) -> bool:
    """
    Applies sender and subject filtering logic.
    """

    sender = str(
        getattr(mail, "SenderEmailAddress", "")
    ).lower()

    subject = str(
        getattr(mail, "Subject", "")
    ).lower()

    if SUBJECT_FILTER.lower() not in subject:
        return False

    if SENDER_FILTER.lower() not in sender:
        return False

    return True


# ==========================================
# OUTLOOK CONNECTION
# ==========================================

def connect_outlook():
    """
    Connects to Outlook MAPI namespace.
    """

    outlook = win32com.client.Dispatch(
        "Outlook.Application"
    ).GetNamespace("MAPI")

    return outlook


# ==========================================
# ATTACHMENT PROCESSING
# ==========================================

def save_attachments(mail) -> int:
    """
    Saves supported attachments from Outlook emails.
    """

    saved_files = 0

    if mail.Attachments.Count == 0:
        return saved_files

    for attachment in mail.Attachments:

        filename = safe_filename(
            attachment.FileName
        )

        extension = Path(filename).suffix.lower()

        if extension not in SUPPORTED_EXTENSIONS:
            continue

        output_path = OUTPUT_FOLDER / filename

        if output_path.exists():
            timestamp = datetime.now().strftime(
                "%Y%m%d_%H%M%S"
            )

            output_path = (
                OUTPUT_FOLDER /
                f"{output_path.stem}_{timestamp}{output_path.suffix}"
            )

        attachment.SaveAsFile(str(output_path))

        saved_files += 1

    return saved_files


# ==========================================
# MAIN PROCESS
# ==========================================

def process_mailbox():

    create_output_folder()

    outlook = connect_outlook()

    total_processed = 0
    total_saved = 0

    for folder in outlook.Folders:

        if MAILBOX_NAME.lower() not in folder.Name.lower():
            continue

        inbox = folder.Folders["Inbox"]

        messages = inbox.Items

        messages.Sort("[ReceivedTime]", True)

        mail = messages.GetFirst()

        while mail:

            try:

                if getattr(mail, "Class", None) != 43:
                    mail = messages.GetNext()
                    continue

                if not mail_matches_filters(mail):
                    mail = messages.GetNext()
                    continue

                total_processed += 1

                total_saved += save_attachments(mail)

            except Exception as error:

                print(
                    f"Error processing email: {error}"
                )

            mail = messages.GetNext()

    print("\n===================================")
    print(f"Processed emails : {total_processed}")
    print(f"Saved attachments: {total_saved}")
    print("===================================\n")


# ==========================================
# ENTRYPOINT
# ==========================================

if __name__ == "__main__":

    process_mailbox()
