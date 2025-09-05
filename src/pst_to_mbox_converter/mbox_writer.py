import mailbox
import logging
from email import message_from_bytes

class MboxWriter:
    """
    A class to handle writing email messages to an MBOX file.
    """
    def __init__(self, mbox_file_path):
        """
        Initializes the MboxWriter and opens the MBOX file.

        Args:
            mbox_file_path (str): The path to the MBOX file.
        """
        try:
            self.mbox = mailbox.mbox(mbox_file_path, create=True)
            logging.info(f"Successfully opened or created MBOX file: {mbox_file_path}")
        except Exception as e:
            logging.error(f"Failed to open MBOX file: {e}")
            raise

    def add_message(self, message_bytes):
        """
        Adds a single message to the MBOX file.

        Args:
            message_bytes (bytes): The raw content of the email message.
        """
        if not message_bytes:
            logging.warning("Attempted to add an empty message. Skipping.")
            return
        try:
            # The mailbox library works with email.message.Message objects.
            # We can create one from the raw bytes.
            # The 'From ' line will be automatically managed by the mailbox library.
            message = message_from_bytes(message_bytes)
            self.mbox.add(message)
        except Exception as e:
            # Log the error but continue, so one bad message doesn't stop the whole process.
            logging.error(f"Failed to parse and add message to MBOX: {e}")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def close(self):
        """
        Closes the MBOX file to ensure all data is written to disk.
        """
        if self.mbox:
            self.mbox.close()
            logging.info("MBOX file has been closed.")
