#!/usr/bin/env python3
"""
MBOX to PDF Converter with HTML Support and Attachment Extraction

Converts mbox email files to PDF format with proper handling of HTML content
and extracts attachments to a separate folder. Shows clean progress display.
"""

import argparse
import email
import logging
import mailbox
import os
import sys
import platform
import ssl
import re
import mimetypes
import time
from email.header import decode_header
from pathlib import Path

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet


# Fix SSL issues on macOS
if platform.system() == 'Darwin':
    ssl._create_default_https_context = ssl._create_unverified_context


def html_to_text(html_content):
    """Convert HTML to plain text by removing tags."""
    if not html_content:
        return ""
        
    # Remove DOCTYPE, html, head tags and their content
    html_content = re.sub(r'<!DOCTYPE[^>]*>|<html[^>]*>|<head>.*?</head>', '', html_content, flags=re.DOTALL)
    
    # Replace common entities
    html_content = html_content.replace('&nbsp;', ' ')
    html_content = html_content.replace('&lt;', '<')
    html_content = html_content.replace('&gt;', '>')
    html_content = html_content.replace('&amp;', '&')
    html_content = html_content.replace('&quot;', '"')
    
    # Replace <br>, <p>, <div> tags with newlines
    html_content = re.sub(r'<br[^>]*>|</p>|</div>', '\n', html_content)
    
    # Remove all other HTML tags
    html_content = re.sub(r'<[^>]+>', ' ', html_content)
    
    # Fix excessive whitespace
    html_content = re.sub(r'\s+', ' ', html_content)
    html_content = re.sub(r'\n\s+', '\n', html_content)
    html_content = re.sub(r'\n+', '\n\n', html_content)
    
    return html_content.strip()


class ProgressTracker:
    """Track and display progress for long-running tasks."""
    
    def __init__(self, total, description="Processing"):
        """Initialize with total number of items to process."""
        self.total = total
        self.current = 0
        self.start_time = time.time()
        self.description = description
        self.last_update_time = 0
        self.update_interval = 0.5  # Update display every half second
        
    def update(self, increment=1):
        """Update progress by the specified increment."""
        self.current += increment
        
        # Only update the display if enough time has passed
        current_time = time.time()
        if current_time - self.last_update_time >= self.update_interval or self.current >= self.total:
            self.last_update_time = current_time
            self.display()
            
    def display(self):
        """Display the current progress."""
        # Calculate percentage
        percent = self.current / self.total * 100 if self.total > 0 else 0
        
        # Calculate elapsed time
        elapsed = time.time() - self.start_time
        
        # Calculate ETA
        if self.current > 0:
            items_per_second = self.current / elapsed
            remaining_items = self.total - self.current
            eta = remaining_items / items_per_second if items_per_second > 0 else 0
        else:
            eta = 0
            
        # Format times
        elapsed_str = self.format_time(elapsed)
        eta_str = self.format_time(eta)
        
        # Create progress bar (width 30 chars)
        bar_width = 30
        filled_width = int(bar_width * self.current / self.total) if self.total > 0 else 0
        bar = "█" * filled_width + "░" * (bar_width - filled_width)
        
        # Build the progress line
        line = f"\r{self.description}: {self.current}/{self.total} [{bar}] {percent:.1f}% | Elapsed: {elapsed_str} | ETA: {eta_str}"
        
        # Ensure the line is at least 80 chars to overwrite previous output
        line = line.ljust(80)
        
        # Print the line (without newline)
        print(line, end="", flush=True)
        
        # Add a newline if we're done
        if self.current >= self.total:
            print()
            
    def format_time(self, seconds):
        """Format time in seconds to a human-readable string."""
        if seconds < 60:
            return f"{seconds:.1f}s"
        elif seconds < 3600:
            minutes = seconds // 60
            seconds = seconds % 60
            return f"{int(minutes)}m {int(seconds)}s"
        else:
            hours = seconds // 3600
            minutes = (seconds % 3600) // 60
            return f"{int(hours)}h {int(minutes)}m"


class MboxConverter:
    """MBOX file to PDF converter with HTML support and attachment extraction."""

    def __init__(self, input_file, output_dir, extract_attachments=True, attachments_dir=None, quiet=False):
        """Initialize the converter."""
        self.input_file = input_file
        self.output_dir = output_dir
        self.extract_attachments = extract_attachments
        self.quiet = quiet
        
        # Set up directories
        os.makedirs(output_dir, exist_ok=True)
        
        # Set up attachments directory
        if extract_attachments:
            if attachments_dir:
                self.attachments_dir = attachments_dir
            else:
                self.attachments_dir = os.path.join(output_dir, "attachments")
            os.makedirs(self.attachments_dir, exist_ok=True)
            
        # Initialize attachment tracking
        self.attachment_count = 0
        self.attachment_saved = 0
        
        self.setup_logging()
        
    def setup_logging(self):
        """Set up a file logger for detailed logs and minimal console output."""
        self.logger = logging.getLogger("mbox2pdf")
        self.logger.setLevel(logging.INFO)
        
        # File handler - detailed logs
        log_file = os.path.join(self.output_dir, "conversion.log")
        file_handler = logging.FileHandler(log_file)
        file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        self.logger.addHandler(file_handler)
        
        # Console handler - errors only if quiet mode is on
        if not self.quiet:
            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.ERROR)  # Only show errors
            console_handler.setFormatter(logging.Formatter("ERROR: %(message)s"))
            self.logger.addHandler(console_handler)
        
    def decode_header_value(self, value):
        """Safely decode email header value."""
        if value is None:
            return ""
            
        result = ""
        for chunk, encoding in decode_header(value):
            if isinstance(chunk, bytes):
                if encoding:
                    try:
                        chunk = chunk.decode(encoding)
                    except:
                        chunk = chunk.decode('utf-8', errors='replace')
                else:
                    chunk = chunk.decode('utf-8', errors='replace')
            result += str(chunk)
        return result
    
    def get_safe_filename(self, filename):
        """Create a safe filename without problematic characters."""
        if not filename:
            return "unknown_file"
            
        # Remove characters that are problematic in filenames
        safe_name = "".join(c if c.isalnum() or c in "-_. " else "_" for c in filename)
        # Limit length
        safe_name = safe_name[:100]
        return safe_name
        
    def extract_email_attachments(self, message, email_id):
        """Extract attachments from an email."""
        attachments = []
        
        if not self.extract_attachments:
            return attachments
            
        # Process each part of the email
        for part in message.walk():
            # Check if this part is an attachment
            content_disposition = part.get("Content-Disposition", "")
            if "attachment" not in content_disposition and "inline" not in content_disposition:
                continue
                
            # Get the filename
            filename = part.get_filename()
            if filename:
                filename = self.decode_header_value(filename)
            else:
                # If no filename, try to create one based on content type
                content_type = part.get_content_type()
                ext = mimetypes.guess_extension(content_type) or ".bin"
                filename = f"attachment_{len(attachments)+1}{ext}"
                
            # Make filename safe
            safe_filename = self.get_safe_filename(filename)
            
            # Create unique filename
            unique_filename = f"{email_id:04d}_{safe_filename}"
            file_path = os.path.join(self.attachments_dir, unique_filename)
            
            # Handle filename collisions
            counter = 1
            while os.path.exists(file_path):
                base_name, ext = os.path.splitext(unique_filename)
                unique_filename = f"{base_name}_{counter}{ext}"
                file_path = os.path.join(self.attachments_dir, unique_filename)
                counter += 1
                
            # Extract attachment content
            try:
                payload = part.get_payload(decode=True)
                if payload:
                    with open(file_path, "wb") as f:
                        f.write(payload)
                        
                    # Record attachment info
                    attachments.append({
                        "filename": filename,
                        "saved_as": unique_filename,
                        "path": file_path,
                        "size": len(payload),
                        "content_type": part.get_content_type()
                    })
                    
                    self.attachment_saved += 1
                    self.logger.info(f"Saved attachment: {filename} as {unique_filename}")
            except Exception as e:
                self.logger.error(f"Error saving attachment {filename}: {e}")
                
        self.attachment_count += len(attachments)
        return attachments
    
    def get_email_text(self, message):
        """Extract text from email message, handling both plain text and HTML."""
        plain_text = ""
        html_text = ""
        
        if message.is_multipart():
            for part in message.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition", ""))
                
                # Skip attachments
                if "attachment" in content_disposition:
                    continue
                    
                try:
                    # Get the content
                    payload = part.get_payload(decode=True)
                    if not payload:
                        continue
                        
                    # Decode with proper charset
                    charset = part.get_content_charset() or 'utf-8'
                    decoded_payload = payload.decode(charset, errors='replace')
                    
                    # Store based on content type
                    if content_type == "text/plain":
                        plain_text += decoded_payload + "\n\n"
                    elif content_type == "text/html":
                        html_text += decoded_payload
                except Exception as e:
                    self.logger.warning(f"Error processing part: {e}")
        else:
            # Single part message
            try:
                payload = message.get_payload(decode=True)
                if payload:
                    charset = message.get_content_charset() or 'utf-8'
                    decoded_payload = payload.decode(charset, errors='replace')
                    
                    content_type = message.get_content_type()
                    if content_type == "text/plain":
                        plain_text = decoded_payload
                    elif content_type == "text/html":
                        html_text = decoded_payload
            except Exception as e:
                self.logger.warning(f"Error processing message: {e}")
        
        # Prefer plain text if available, otherwise convert HTML to text
        if plain_text.strip():
            return plain_text
        elif html_text:
            return html_to_text(html_text)
        else:
            return "[No message content found]"
    
    def format_size(self, size_bytes):
        """Format file size in human-readable format."""
        if size_bytes < 1024:
            return f"{size_bytes} bytes"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes/1024:.1f} KB"
        else:
            return f"{size_bytes/(1024*1024):.1f} MB"
    
    def create_pdf(self, message, output_path, email_id):
        """Create a PDF file from an email message."""
        try:
            # Extract email metadata
            subject = self.decode_header_value(message.get("Subject", "No Subject"))
            from_addr = self.decode_header_value(message.get("From", "Unknown"))
            to_addr = self.decode_header_value(message.get("To", "Unknown"))
            date = message.get("Date", "Unknown Date")
            
            # Extract attachments
            attachments = self.extract_email_attachments(message, email_id) if self.extract_attachments else []
            
            # Get email body
            body_text = self.get_email_text(message)
            
            # Create PDF
            doc = SimpleDocTemplate(output_path, pagesize=letter)
            styles = getSampleStyleSheet()
            
            # Prepare content
            content = []
            
            # Add headers
            content.append(Paragraph(f"<b>Subject:</b> {subject}", styles['Heading2']))
            content.append(Paragraph(f"<b>From:</b> {from_addr}", styles['Normal']))
            content.append(Paragraph(f"<b>To:</b> {to_addr}", styles['Normal']))
            content.append(Paragraph(f"<b>Date:</b> {date}", styles['Normal']))
            
            # Add attachments section if any
            if attachments:
                content.append(Spacer(1, 12))
                content.append(Paragraph(f"<b>Attachments:</b>", styles['Heading3']))
                
                # Create a table of attachments
                attachment_data = []
                for i, attachment in enumerate(attachments):
                    filename = attachment["filename"]
                    size = self.format_size(attachment["size"])
                    saved_as = attachment["saved_as"]
                    attachment_data.append([f"{i+1}.", filename, size, saved_as])
                
                if attachment_data:
                    table = Table(attachment_data, colWidths=[20, 200, 70, 200])
                    table.setStyle(TableStyle([
                        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 0), (-1, -1), 9),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                        ('TOPPADDING', (0, 0), (-1, -1), 2),
                        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ]))
                    content.append(table)
            
            content.append(Spacer(1, 12))
            
            # Add a separator line
            content.append(Paragraph("<hr/>", styles['Normal']))
            
            content.append(Spacer(1, 12))
            
            # Add body
            if body_text:
                # Process the text to handle ReportLab's XML parsing requirements
                # Replace special characters that might cause XML parsing errors
                body_text = body_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                
                # Split into paragraphs
                for paragraph in body_text.split('\n\n'):
                    if paragraph.strip():
                        # Replace single newlines with <br/> tags
                        formatted_text = paragraph.replace('\n', '<br/>')
                        try:
                            p = Paragraph(formatted_text, styles['Normal'])
                            content.append(p)
                            content.append(Spacer(1, 6))
                        except Exception as e:
                            # If paragraph fails, try a simplified version
                            self.logger.warning(f"Error formatting paragraph: {e}")
                            try:
                                simple_text = "".join(c if c.isascii() else '_' for c in paragraph)
                                content.append(Paragraph(simple_text, styles['Normal']))
                                content.append(Spacer(1, 6))
                            except:
                                # Last resort
                                content.append(Paragraph("[Formatting error with this paragraph]", styles['Normal']))
                                content.append(Spacer(1, 6))
            else:
                content.append(Paragraph("[No message body]", styles['Normal']))
            
            # Build PDF
            doc.build(content)
            return True
            
        except Exception as e:
            self.logger.error(f"Error creating PDF: {e}")
            return False
    
    def process_mbox(self):
        """Process all messages in the mbox file."""
        try:
            # Open the mbox file
            self.logger.info(f"Opening mbox file: {self.input_file}")
            mbox = mailbox.mbox(self.input_file)
            
            # Process each message
            total = len(mbox)
            success = 0
            
            self.logger.info(f"Found {total} messages in the mbox file")
            
            # Initialize progress tracker
            progress = ProgressTracker(total, "Converting emails")
            
            # Start time for overall stats
            start_time = time.time()
            
            for i, message in enumerate(mbox):
                try:
                    # Get subject for filename
                    subject = self.decode_header_value(message.get("Subject", "No Subject"))
                    safe_subject = "".join(c if c.isalnum() or c in "-_ " else "_" for c in subject)
                    safe_subject = safe_subject[:50]  # Limit length
                    
                    # Create output filename
                    filename = f"{i+1:04d}_{safe_subject}.pdf"
                    output_path = os.path.join(self.output_dir, filename)
                    
                    # Create PDF
                    if self.create_pdf(message, output_path, i+1):
                        success += 1
                except Exception as e:
                    self.logger.error(f"Error processing message {i+1}: {e}")
                
                # Update progress
                progress.update()
            
            # Calculate total time
            total_time = time.time() - start_time
            time_per_email = total_time / total if total > 0 else 0
            
            # Print summary
            print("\nConversion Summary:")
            print(f"  Processed {total} emails in {progress.format_time(total_time)} ({time_per_email:.2f} seconds per email)")
            print(f"  Successfully converted: {success} ({success/total*100:.1f}%)")
            print(f"  Failed: {total-success}")
            
            if self.extract_attachments:
                print(f"  Attachments extracted: {self.attachment_saved} (from {self.attachment_count} found)")
                print(f"  Attachments directory: {os.path.abspath(self.attachments_dir)}")
            
            print(f"  Output directory: {os.path.abspath(self.output_dir)}")
            print(f"  Detailed log: {os.path.join(os.path.abspath(self.output_dir), 'conversion.log')}")
            
            # Log summary
            self.logger.info(f"Conversion complete. Processed {total} emails in {total_time:.1f} seconds")
            self.logger.info(f"Successfully converted {success} of {total} emails ({success/total*100:.1f}%)")
            if self.extract_attachments:
                self.logger.info(f"Extracted {self.attachment_saved} of {self.attachment_count} attachments")
            
            return success, total
            
        except Exception as e:
            self.logger.error(f"Error processing mbox file: {e}")
            return 0, 0


def main():
    """Main function."""
    # Parse arguments
    parser = argparse.ArgumentParser(description="Convert mbox file to PDFs with clean progress display")
    parser.add_argument("input", help="Input mbox file")
    parser.add_argument("output", help="Output directory for PDFs")
    parser.add_argument("--no-attachments", action="store_true", help="Skip extracting attachments")
    parser.add_argument("--attachments-dir", help="Custom directory for attachments (default: output_dir/attachments)")
    parser.add_argument("--quiet", action="store_true", help="Minimize console output, showing only progress and summary")
    args = parser.parse_args()
    
    # Check if input file exists
    if not os.path.isfile(args.input):
        print(f"Error: Input file '{args.input}' not found")
        return 1
    
    # Convert mbox to PDFs
    converter = MboxConverter(
        args.input, 
        args.output, 
        extract_attachments=not args.no_attachments,
        attachments_dir=args.attachments_dir,
        quiet=args.quiet or True  # Always use quiet mode for clean output
    )
    success, total = converter.process_mbox()
    
    # Return success if most messages were converted
    return 0 if success > total * 0.9 else 1


if __name__ == "__main__":
    sys.exit(main())
