import ctypes
import ctypes.util
import logging
import platform

class LibPffError(Exception):
    """Custom exception for libpff errors."""
    pass

# --- ctypes definitions for libpff ---

# Find and load the shared library
lib_name = 'pff'
if platform.system() == "Windows":
    lib_path = ctypes.util.find_library('libpff.dll') or ctypes.util.find_library('pff.dll')
else:
    lib_path = ctypes.util.find_library(lib_name)

if not lib_path:
    lib_path = ctypes.util.find_library('libpff')
    if not lib_path:
        raise ImportError("Could not find libpff shared library.")

try:
    libpff = ctypes.CDLL(lib_path)
except OSError as e:
    raise ImportError(f"Failed to load libpff shared library from {lib_path}: {e}")

# Opaque pointers for libpff types
pff_file_t = ctypes.c_void_p
pff_item_t = ctypes.c_void_p
pff_folder_t = pff_item_t
pff_message_t = pff_item_t

# MAPI Property IDs
PR_TRANSPORT_MESSAGE_HEADERS = 0x007d001e
PR_BODY = 0x1000001e
PR_SUBJECT = 0x0037001e

# --- Function Signatures ---

# File functions
libpff.libpff_file_initialize.restype = ctypes.c_int
libpff.libpff_file_initialize.argtypes = [ctypes.POINTER(pff_file_t), ctypes.c_void_p]
libpff.libpff_file_open.restype = ctypes.c_int
libpff.libpff_file_open.argtypes = [pff_file_t, ctypes.c_char_p, ctypes.c_int]
libpff.libpff_file_close.restype = ctypes.c_int
libpff.libpff_file_close.argtypes = [pff_file_t]
libpff.libpff_file_free.restype = ctypes.c_int
libpff.libpff_file_free.argtypes = [pff_file_t]
libpff.libpff_file_get_root_folder.restype = ctypes.c_int
libpff.libpff_file_get_root_folder.argtypes = [pff_file_t, ctypes.POINTER(pff_folder_t)]

# Folder functions
libpff.libpff_folder_get_number_of_sub_folders.restype = ctypes.c_int
libpff.libpff_folder_get_number_of_sub_folders.argtypes = [pff_folder_t, ctypes.POINTER(ctypes.c_int)]
libpff.libpff_folder_get_sub_folder.restype = ctypes.c_int
libpff.libpff_folder_get_sub_folder.argtypes = [pff_folder_t, ctypes.c_int, ctypes.POINTER(pff_folder_t)]
libpff.libpff_folder_get_number_of_sub_messages.restype = ctypes.c_int
libpff.libpff_folder_get_number_of_sub_messages.argtypes = [pff_folder_t, ctypes.POINTER(ctypes.c_int)]
libpff.libpff_folder_get_sub_message.restype = ctypes.c_int
libpff.libpff_folder_get_sub_message.argtypes = [pff_folder_t, ctypes.c_int, ctypes.POINTER(pff_message_t)]
libpff.libpff_folder_get_utf8_name_size.restype = ctypes.c_int
libpff.libpff_folder_get_utf8_name_size.argtypes = [pff_folder_t, ctypes.POINTER(ctypes.c_size_t)]
libpff.libpff_folder_get_utf8_name.restype = ctypes.c_int
libpff.libpff_folder_get_utf8_name.argtypes = [pff_folder_t, ctypes.c_char_p, ctypes.c_size_t]

# Message functions
libpff.libpff_message_get_entry_value_utf8_string_size.restype = ctypes.c_int
libpff.libpff_message_get_entry_value_utf8_string_size.argtypes = [pff_message_t, ctypes.c_uint32, ctypes.POINTER(ctypes.c_size_t)]
libpff.libpff_message_get_entry_value_utf8_string.restype = ctypes.c_int
libpff.libpff_message_get_entry_value_utf8_string.argtypes = [pff_message_t, ctypes.c_uint32, ctypes.c_char_p, ctypes.c_size_t]

# --- PSTReader implementation ---

class PSTReader:
    def __init__(self, pst_file_path):
        self.pst_file = pff_file_t()
        self.root_folder = None
        if libpff.libpff_file_initialize(ctypes.byref(self.pst_file), None) != 1:
            raise LibPffError("Failed to initialize libpff file object.")

        try:
            if libpff.libpff_file_open(self.pst_file, pst_file_path.encode('utf-8'), 1) != 1:
                libpff.libpff_file_free(self.pst_file)
                self.pst_file = None
                raise LibPffError(f"Failed to open PST file: {pst_file_path}")
        except LibPffError:
            raise

        self.root_folder = pff_folder_t()
        if libpff.libpff_file_get_root_folder(self.pst_file, ctypes.byref(self.root_folder)) != 1:
            self.close()
            raise LibPffError("Failed to get root folder from PST file.")

        logging.info(f"Successfully opened PST file: {pst_file_path}")

    def get_messages(self):
        logging.info("Starting to extract messages from the PST file.")
        message_count = 0
        if self.root_folder:
            for message_bytes in self._folder_iterator(self.root_folder):
                yield message_bytes
                message_count += 1
                if message_count > 0 and message_count % 100 == 0:
                    logging.info(f"Extracted {message_count} messages so far...")
        logging.info(f"Finished extracting. Total messages found: {message_count}")

    def _get_folder_name(self, folder):
        name_size = ctypes.c_size_t(0)
        if libpff.libpff_folder_get_utf8_name_size(folder, ctypes.byref(name_size)) != 1:
            return "(Unknown Folder)"

        name_buffer = ctypes.create_string_buffer(name_size.value + 1)
        if libpff.libpff_folder_get_utf8_name(folder, name_buffer, name_size.value + 1) != 1:
            return "(Unknown Folder)"

        return name_buffer.value.decode('utf-8', errors='ignore')

    def _folder_iterator(self, folder):
        folder_name = self._get_folder_name(folder)
        num_messages = ctypes.c_int(0)
        if libpff.libpff_folder_get_number_of_sub_messages(folder, ctypes.byref(num_messages)) == 1 and num_messages.value > 0:
            logging.info(f"Scanning folder: {folder_name} ({num_messages.value} messages)")
            for i in range(num_messages.value):
                message = pff_message_t()
                if libpff.libpff_folder_get_sub_message(folder, i, ctypes.byref(message)) == 1:
                    msg_bytes = self._get_message_as_bytes(message)
                    if msg_bytes:
                        yield msg_bytes

        num_sub_folders = ctypes.c_int(0)
        if libpff.libpff_folder_get_number_of_sub_folders(folder, ctypes.byref(num_sub_folders)) == 1 and num_sub_folders.value > 0:
            for i in range(num_sub_folders.value):
                sub_folder = pff_folder_t()
                if libpff.libpff_folder_get_sub_folder(folder, i, ctypes.byref(sub_folder)) == 1:
                    yield from self._folder_iterator(sub_folder)

    def _get_entry_value_string(self, item_pointer, entry_type):
        size = ctypes.c_size_t(0)
        if libpff.libpff_message_get_entry_value_utf8_string_size(item_pointer, entry_type, ctypes.byref(size)) != 1 or size.value == 0:
            return None

        buffer = ctypes.create_string_buffer(size.value)
        if libpff.libpff_message_get_entry_value_utf8_string(item_pointer, entry_type, buffer, size.value) != 1:
            return None

        return buffer.value

    def _get_message_as_bytes(self, message):
        try:
            headers_bytes = self._get_entry_value_string(message, PR_TRANSPORT_MESSAGE_HEADERS)

            if not headers_bytes:
                subject_bytes = self._get_entry_value_string(message, PR_SUBJECT)
                if subject_bytes:
                    headers_bytes = b"Subject: " + subject_bytes
                else:
                    logging.warning("Message has no transport headers or subject. Constructing a minimal header.")
                    headers_bytes = b"Subject: (No Subject)"

            body_bytes = self._get_entry_value_string(message, PR_BODY) or b''

            if not headers_bytes.endswith(b"\r\n\r\n"):
                headers_bytes = headers_bytes.strip() + b"\r\n\r\n"

            return headers_bytes + body_bytes

        except Exception as e:
            logging.error(f"Error reading message bytes: {e}")
            return None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def close(self):
        if hasattr(self, 'pst_file') and self.pst_file:
            libpff.libpff_file_close(self.pst_file)
            self.pst_file = None
            self.root_folder = None
        logging.info("PST file closed.")
