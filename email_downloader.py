import os
import comtypes.client


def download_attachments(email, download_folder, extension=".xlsx"):
    """
    Download attachments with a given extension from a single Outlook email.
    """
    if email.Attachments.Count == 0:
        return

    for i in range(1, email.Attachments.Count + 1):
        try:
            attachment = email.Attachments.Item(i)
            filename = attachment.FileName

            if filename.lower().endswith(extension):
                save_path = os.path.join(download_folder, filename)
                attachment.SaveAsFile(save_path)
        except Exception as e:
            print(f"Error downloading attachment {i}: {e}")


def find_emails(senders, subject_keyword, download_folder, max_emails=50, debug=False):
    """
    Search recent emails in the default inbox, filter by senders + subject keyword,
    and download .xlsx attachments for matches.

    Parameters
    ----------
    senders : list of str
        One or more substrings to match in SenderEmailAddress.
    subject_keyword : str
        Substring to look for in email.Subject.
    download_folder : str
        Local folder path where attachments will be saved.
    max_emails : int
        Maximum number of recent emails to search.
    debug : bool, optional (default=False)
        If True, prints email metadata (subject, sender) for inspection.
    """
    try:
        outlook = comtypes.client.CreateObject("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)

        recent_messages = []
        message = messages.GetFirst()
        count = 0

        # Collect recent N messages
        while message and count < max_emails:
            recent_messages.append(message)
            count += 1
            message = messages.GetNext()

        # Inspect + process
        for idx, msg in enumerate(recent_messages, start=1):

            if debug:
                print(f"Email {idx}:")
                print(f"Subject: {msg.Subject}")
                print(f"Sender:  {msg.SenderEmailAddress}")
                print("-" * 40)

            # Filtering logic
            sender_match = any(
                sender.lower() in (msg.SenderEmailAddress or "").lower()
                for sender in senders
            )
            subject_match = subject_keyword.lower() in (msg.Subject or "").lower()

            if sender_match and subject_match:
                download_attachments(msg, download_folder)

    except Exception as e:
        print(f"Error occurred while searching emails: {e}")
