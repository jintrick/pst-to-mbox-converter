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
        try:
            headers = message.get_transport_headers()
            body = message.get_plain_text_body()

            if not headers:
                logging.warning(f"Message with subject '{message.subject}' has no transport headers. Constructing a minimal header.")
                subject = message.subject or "(No Subject)"
                headers = f"Subject: {subject}\r\n"

            if isinstance(headers, str):
                headers_bytes = headers.encode("utf-8", errors="replace")
            else:
                headers_bytes = headers if headers else b""

            if isinstance(body, str):
                body_bytes = body.encode("utf-8", errors="replace")
            else:
                body_bytes = body if body else b""

            if not headers_bytes.endswith(b"\r\n\r\n"):
                headers_bytes = headers_bytes.strip() + b"\r\n\r\n"

            return headers_bytes + body_bytes

        except Exception as e:
            logging.error(f"Error reading message bytes for subject '{message.subject}': {e}")
            return None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def close(self):
        self.pst_file.close()
        logging.info("PST file closed.")
