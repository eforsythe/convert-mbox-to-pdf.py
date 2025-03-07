# MBOX to PDF Converter

A powerful Python tool to convert MBOX email archives to PDF files with proper formatting, and attachment extraction.

## Features

- **HTML Email Support**: Properly handles HTML-formatted emails while preserving readable content
- **Attachment Extraction**: Automatically saves email attachments to a separate folder
- **Clean Progress Display**: Visual progress bar with time estimates
- **Proper Encoding**: Handles special characters and international text correctly
- **Well-Formatted PDFs**: Creates readable, well-structured PDF documents

## Installation

1. Install the required dependencies:
   ```bash
   pip install reportlab tqdm
   ```

## Usage

### Converting an MBOX File

```bash
python convert-mbox-to-pdf.py input.mbox output_directory
```

### Command Line Options

```
Single File Mode:
  --no-attachments       Skip extracting attachments
  --attachments-dir DIR  Custom directory for attachments
  --quiet                Minimize console output

## Examples

### Convert a Single MBOX File

```bash
python convert-mbox-to-pdf.py ~/mail/inbox.mbox ~/Documents/mail_backup
```

### Skip Attachment Extraction

```bash
python mbox_converter.py inbox.mbox output_dir --no-attachments
```

### Specify Custom Attachment Directory

```bash
python mbox_converter.py inbox.mbox output_dir --attachments-dir ~/Documents/email_attachments
```

## Output Structure

```
output_directory/
├── 0001_first_email_subject.pdf
├── 0002_second_email_subject.pdf
├── ...
├── attachments/
│   ├── 0001_attachment1.docx
│   ├── 0005_attachment2.jpg
│   └── ...
└── conversion.log
```

## Each PDF Contains

- Email headers (From, To, Subject, Date)
- List of attachments with filenames and sizes
- Complete email body text

## Requirements

- Python 3.6+
- reportlab 3.6+ (for PDF generation)
- tqdm (for progress display)

## Limitations

- Very large attachments might be skipped (configurable size limit)
- Some complex HTML formatting may be simplified
- Special email features like encrypted content are not supported

## License

See the LICENSE file for details.

## Acknowledgments

- ReportLab for the PDF generation library
- The Python mailbox library for MBOX parsing

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
