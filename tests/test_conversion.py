import pytest
import os
import mailbox
import sys
import email
from pst_to_mbox_converter.pst_reader import PSTReader
from pst_to_mbox_converter.mbox_writer import MboxWriter
from pst_to_mbox_converter import main as main_module

SAMPLE_PST_PATH = "tests/data/sample.pst"
# A sample valid email message in bytes
VALID_MESSAGE_BYTES = b"From MAILER-DAEMON Fri Jul  8 12:08:34 2022\nSubject: Test\n\nThis is a test."

def test_pst_reader_opens_file():
    """Tests that PSTReader can successfully open a PST file."""
    try:
        reader = PSTReader(SAMPLE_PST_PATH)
        reader.close()
    except Exception as e:
        pytest.fail(f"PSTReader failed to open the PST file: {e}")

def test_pst_reader_get_messages_yields_bytes():
    """Tests that get_messages yields non-empty bytes."""
    reader = PSTReader(SAMPLE_PST_PATH)
    messages = list(reader.get_messages())
    reader.close()

    assert len(messages) > 0, "No messages were extracted from the PST file."
    for msg in messages:
        assert isinstance(msg, bytes), f"Message should be bytes, but got {type(msg)}."
        assert len(msg) > 0, "Message bytes should not be empty."

def test_mbox_writer_adds_message(tmp_path):
    """Tests that MboxWriter can add a message to an MBOX file."""
    mbox_file = tmp_path / "test.mbox"
    writer = MboxWriter(str(mbox_file))
    writer.add_message(VALID_MESSAGE_BYTES)
    writer.close()

    assert os.path.exists(mbox_file)
    with open(mbox_file, "r") as f:
        content = f.read()
        assert "Subject: Test" in content

def test_mbox_writer_handles_none_message(tmp_path, caplog):
    """Tests that MboxWriter logs a warning for None messages."""
    mbox_file = tmp_path / "test.mbox"
    writer = MboxWriter(str(mbox_file))
    writer.add_message(None)
    writer.close()

    assert "Attempted to add an empty message. Skipping." in caplog.text
    assert any(record.levelno == 30 for record in caplog.records) # 30 is WARNING

def test_mbox_writer_handles_corrupted_message(tmp_path, caplog, monkeypatch):
    """Tests that MboxWriter logs an error for corrupted messages."""

    def mock_message_from_bytes(b, *, policy=None):
        raise ValueError("Simulated parsing error")

    monkeypatch.setattr("pst_to_mbox_converter.mbox_writer.message_from_bytes", mock_message_from_bytes)

    mbox_file = tmp_path / "test.mbox"
    writer = MboxWriter(str(mbox_file))
    writer.add_message(b"any bytes will now cause an error")
    writer.close()

    assert "Failed to parse and add message to MBOX" in caplog.text
    assert "Simulated parsing error" in caplog.text
    assert any(record.levelno == 40 for record in caplog.records) # 40 is ERROR

def test_end_to_end_conversion(tmp_path, monkeypatch):
    """
    Tests the full conversion process from a sample PST to an MBOX file.
    """
    output_mbox = tmp_path / "output.mbox"

    # Mock sys.argv
    monkeypatch.setattr(
        "sys.argv",
        ["pst-converter", "--input", SAMPLE_PST_PATH, "--output", str(output_mbox)]
    )

    # I'll mock sys.exit to prevent the test runner from exiting
    def mock_exit(status):
        pass
    monkeypatch.setattr(sys, "exit", mock_exit)

    main_module.main()

    assert output_mbox.exists()

    # Check if the mbox file is not empty and contains more than one message
    assert output_mbox.stat().st_size > 0

    # For a more thorough check, we can use the mailbox library to count messages
    mbox = mailbox.mbox(output_mbox)
    assert len(mbox) > 0
