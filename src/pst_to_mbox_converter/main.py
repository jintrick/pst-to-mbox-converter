import argparse
import logging
import sys
from .pst_reader import PSTReader
from .mbox_writer import MboxWriter

def main():
    """
    The main function for the pst-to-mbox-converter tool.
    It parses command-line arguments, and orchestrates the conversion process.
    """
    parser = argparse.ArgumentParser(
        description="A command-line tool to convert Microsoft Outlook PST files to MBOX format."
    )
    parser.add_argument(
        "-i", "--input",
        required=True,
        help="The path to the input PST file.",
        metavar="PST_FILE"
    )
    parser.add_argument(
        "-o", "--output",
        required=True,
        help="The path for the output MBOX file.",
        metavar="MBOX_FILE"
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose logging for detailed progress."
    )

    args = parser.parse_args()

    # Configure logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(levelname)s - %(message)s',
        stream=sys.stdout
    )

    pst_reader = None
    mbox_writer = None

    try:
        logging.info(f"Starting conversion from {args.input} to {args.output}")

        pst_reader = PSTReader(args.input)
        mbox_writer = MboxWriter(args.output)

        # The core conversion loop
        for message_bytes in pst_reader.get_messages():
            mbox_writer.add_message(message_bytes)

        logging.info("Conversion process finished successfully.")

    except FileNotFoundError as e:
        logging.error(f"Error: Input file not found - {e}")
        sys.exit(1)
    except Exception as e:
        logging.error(f"A critical error occurred during the conversion process: {e}")
        sys.exit(1)
    finally:
        # Ensure resources are closed properly
        if pst_reader:
            pst_reader.close()
        if mbox_writer:
            mbox_writer.close()
        logging.info("Clean-up complete. Exiting.")

if __name__ == "__main__":
    main()
