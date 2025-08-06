"""
This module is able to extract the attachments out of outlook .msg files within a directory.
"""

from tkinter import filedialog
import extract_msg
from pathlib import Path

def extract_attachments(emails: Path, save_To: Path) -> None:
    """
    Extracts attachments from emails folder and save them in a given folder.
    """
    for file_path in emails.iterdir():
        msg = extract_msg.Message(file_path)
        attachments = msg.attachments

        for attachment in attachments:
            attachment.save(customPath=save_To)

if __name__ == '__main__':
    print("Please select the folder containing outlook messages:")
    emails = Path(filedialog.askdirectory())
    print("Please select the folder to save the attachments to:")
    save_to = Path(filedialog.askdirectory())
    extract_attachments(emails, save_to)