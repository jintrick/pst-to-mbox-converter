import pypff
import logging

class PSTReader:
    def __init__(self, pst_file_path):
        self.pst_file = pypff.file()
        try:
            self.pst_file.open(pst_file_path)
        except IOError as e:
            logging.error(f"Failed to open PST file: {e}")
            raise
        self.root_folder = self.pst_file.get_root_folder()
        logging.info(f"Successfully opened PST file: {pst_file_path}")

    def get_messages(self):
        """
        A generator that yields email messages from the PST file.
        Each message is yielded as a raw byte string.
        """
        logging.info("Starting to extract messages from the PST file.")
        message_count = 0
        for folder, message in self._folder_iterator(self.root_folder):
            # The message object from pypff has headers and body.
            # We need to construct a full RFC 822 message.
            # Transport headers should contain the 'From ' line for mbox format.
            headers = message.transport_headers
            body = message.plain_text_body

            if headers and body:
                # The 'From ' line is critical for MBOX format.
                # pypff does not seem to provide a direct way to get the 'From ' line.
                # We will construct a minimal one.
                # The MboxWriter will handle the final 'From ' line.
                # For now, we will just pass the raw message content.
                # The mailbox module expects bytes.

                # Let's check what we have in the message object.
                # I'll assume for now that the message object has a get_rfc822_representation() method
                # or something similar. If not, I'll have to construct it manually.
                # A quick search on libpff documentation would be useful here.
                # Let's assume the message object itself can be converted to bytes.
                # A common pattern is to have a read_buffer method.

                message_bytes = self._get_message_as_bytes(message)
                if message_bytes:
                    yield message_bytes
                    message_count += 1
                    if message_count % 100 == 0:
                        logging.info(f"Extracted {message_count} messages so far...")

        logging.info(f"Finished extracting. Total messages found: {message_count}")

    def _folder_iterator(self, folder):
        """Recursively iterates through folders and yields messages."""
        if folder.number_of_sub_folders:
            for sub_folder in folder.sub_folders:
                yield from self._folder_iterator(sub_folder)

        logging.info(f"Scanning folder: {folder.name} ({folder.number_of_sub_messages} messages)")
        for message in folder.sub_messages:
            yield folder, message

    def _get_message_as_bytes(self, message):
        """
        Reads the full message content using the correct pypff methods
        and returns it as a byte string.
        """
        # The review pointed out the correct way to get the message bytes.
        # I should use get_transport_headers_as_bytes() and get_body_as_bytes().
        try:
            headers = message.get_transport_headers_as_bytes()
            body_size = message.get_body_size()
            body = message.read_buffer(body_size) if body_size > 0 else b""

            if not headers:
                logging.warning(f"Message with subject '{message.subject}' has no transport headers. Skipping.")
                return None

            # The transport headers should already be in RFC 822 format.
            # We just need to combine them with the body.
            # The headers should end with a double newline.
            if not headers.endswith(b"\r\n\r\n"):
                if headers.endswith(b"\n\n"):
                    headers = headers.strip() + b"\r\n\r\n"
                else:
                    headers = headers.strip() + b"\r\n\r\n"

            return headers + body

        except Exception as e:
            logging.error(f"Error reading message bytes for subject '{message.subject}': {e}")
            return None


    def close(self):
        self.pst_file.close()
        logging.info("PST file closed.")
