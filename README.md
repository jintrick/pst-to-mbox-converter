# PST to MBOX Converter

A simple and efficient command-line tool for converting Microsoft Outlook PST files to the standard MBOX format.

## Features

- Converts PST files to MBOX format.
- Recursively processes all folders within the PST file.
- Memory-efficient design for handling large PST files.
- Provides progress logging during conversion.

## Installation

1.  Clone this repository to your local machine.
2.  Navigate to the root directory of the project (`pst-to-mbox-converter`).
3.  Install the tool and its dependencies using pip:

    ```bash
    pip install .
    ```

## Usage

After installation, you can use the `pst-converter` command in your terminal.

### Syntax

```bash
pst-converter --input <path_to_pst_file> --output <path_to_mbox_file>
```

### Arguments

-   `--input`: (Required) The path to the input PST file you want to convert.
-   `--output`: (Required) The path where the output MBOX file will be saved.
-   `--verbose`: (Optional) Enable verbose logging for more detailed output.

### Example

```bash
pst-converter --input /path/to/your/archive.pst --output /path/to/converted/archive.mbox
```

This command will read `archive.pst`, convert its contents, and save them to `archive.mbox`.
