"""
email_sender.py
───────────────
Automatically or manually send Excel files (or other attachments) via email using SMTP.

SETUP (one‑time)
    1.  pip install python-dotenv
    2.  cp .env.example .env          # fill in your SMTP credentials

Gmail users:
    • Enable 2‑Factor Authentication on your Google account.
    • Generate an App Password at https://myaccount.google.com/apppasswords
    • Put that App Password (not your login password) into SMTP_PASS in .env

USAGE:
    # Auto-detect most recent Excel file
    python email_sender.py
    
    # Specify a file manually
    python email_sender.py myfile.xlsx
    python email_sender.py output/coingecko_all_data_20240203_143022.xlsx
    python email_sender.py "path with spaces/report.xlsx"
    
    # Specify custom recipient
    python email_sender.py myfile.xlsx recipient@example.com
"""

import os
import sys
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from datetime import datetime

from dotenv import load_dotenv


def send_email_with_attachment(file_path: Path, recipient: str = None, subject: str = None, body: str = None) -> bool:
    """
    Send an email with the specified file as an attachment.
    
    Args:
        file_path: Path to the file to attach
        recipient: Email address to send to (uses .env RECIPIENT_EMAIL if not provided)
        subject: Email subject line (auto-generated if not provided)
        body: Email body text (auto-generated if not provided)
    
    Returns:
        True if email sent successfully, False otherwise
    """
    # Load environment variables
    load_dotenv()
    
    smtp_host = os.getenv("SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER", "")
    smtp_pass = os.getenv("SMTP_PASS", "")
    
    # Use provided recipient or fall back to .env
    if recipient is None:
        recipient = os.getenv("RECIPIENT_EMAIL", "")
    
    # Validate credentials
    if not all([smtp_user, smtp_pass, recipient]):
        print("\n❌ ERROR: Missing email credentials!")
        print("   Please set the following in your .env file:")
        if not smtp_user:
            print("   - SMTP_USER (your email address)")
        if not smtp_pass:
            print("   - SMTP_PASS (your app password)")
        if not recipient:
            print("   - RECIPIENT_EMAIL (recipient's email address)")
        print("\n   For Gmail users:")
        print("   1. Enable 2-Factor Authentication")
        print("   2. Generate App Password at: https://myaccount.google.com/apppasswords")
        print("   3. Use the App Password (not your login password) for SMTP_PASS\n")
        return False
    
    # Validate file exists
    if not file_path.exists():
        print(f"\n❌ ERROR: File not found: {file_path}\n")
        return False
    
    # Get file info
    file_size_mb = file_path.stat().st_size / (1024 * 1024)
    file_name = file_path.name
    
    print(f"\n📧 Preparing to send email...")
    print(f"   From:       {smtp_user}")
    print(f"   To:         {recipient}")
    print(f"   Attachment: {file_name} ({file_size_mb:.2f} MB)")
    
    # Auto-generate subject if not provided
    if subject is None:
        subject = f"File: {file_name} - {datetime.now().strftime('%d %b %Y')}"
    
    # Auto-generate body if not provided
    if body is None:
        body = f"""Hello,

Please find attached the file: {file_name}

File details:
• Name: {file_name}
• Size: {file_size_mb:.2f} MB
• Sent: {datetime.now().strftime('%d %b %Y at %H:%M')}

Best regards,
Automated Email Sender
"""
    
    # Build MIME message
    msg = MIMEMultipart()
    msg["From"] = smtp_user
    msg["To"] = recipient
    msg["Subject"] = subject
    
    # Attach body text
    msg.attach(MIMEText(body, "plain"))
    
    # Attach file
    try:
        with open(file_path, "rb") as f:
            attachment = MIMEApplication(f.read(), Name=file_name)
        attachment["Content-Disposition"] = f'attachment; filename="{file_name}"'
        msg.attach(attachment)
    except Exception as e:
        print(f"\n❌ ERROR: Failed to read file: {e}\n")
        return False
    
    # Send email
    try:
        print(f"   Connecting to {smtp_host}:{smtp_port}...")
        with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            print(f"   Logging in as {smtp_user}...")
            server.login(smtp_user, smtp_pass)
            print(f"   Sending email...")
            server.sendmail(smtp_user, recipient, msg.as_string())
        
        print(f"\n✅ SUCCESS! Email sent to {recipient}\n")
        return True
        
    except smtplib.SMTPAuthenticationError:
        print("\n❌ ERROR: Authentication failed!")
        print("   For Gmail users:")
        print("   1. Make sure you're using an App Password, not your regular password")
        print("   2. Generate one at: https://myaccount.google.com/apppasswords")
        print("   3. Enable 2-Factor Authentication first if you haven't\n")
        return False
        
    except smtplib.SMTPException as e:
        print(f"\n❌ ERROR: SMTP error occurred: {e}\n")
        return False
        
    except Exception as e:
        print(f"\n❌ ERROR: Failed to send email: {e}\n")
        return False


def find_most_recent_excel() -> Path:
    """
    Automatically find the most recent Excel file in common locations.
    Searches in: current directory, output/, and parent directory.
    
    Returns:
        Path to the most recent .xlsx file found
    
    Raises:
        FileNotFoundError if no Excel files are found
    """
    search_paths = [
        Path("."),                    # Current directory
        Path("output"),               # output folder
        Path("../output"),            # output folder in parent
        Path(".."),                   # Parent directory
    ]
    
    excel_files = []
    
    # Search for Excel files in all paths
    for search_path in search_paths:
        if search_path.exists() and search_path.is_dir():
            # Find all .xlsx files (but not temporary Excel files starting with ~)
            found_files = [
                f for f in search_path.glob("*.xlsx") 
                if not f.name.startswith("~")
            ]
            excel_files.extend(found_files)
    
    if not excel_files:
        raise FileNotFoundError("No Excel files found in current directory or output/ folder")
    
    # Sort by modification time (most recent first)
    excel_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    
    return excel_files[0]


def list_available_excel_files() -> list[Path]:
    """
    List all Excel files found in common locations.
    
    Returns:
        List of Path objects for all .xlsx files found
    """
    search_paths = [
        Path("."),
        Path("output"),
        Path("../output"),
        Path(".."),
    ]
    
    excel_files = []
    
    for search_path in search_paths:
        if search_path.exists() and search_path.is_dir():
            found_files = [
                f for f in search_path.glob("*.xlsx") 
                if not f.name.startswith("~")
            ]
            excel_files.extend(found_files)
    
    # Remove duplicates and sort by modification time
    seen = set()
    unique_files = []
    for f in excel_files:
        abs_path = f.resolve()
        if abs_path not in seen:
            seen.add(abs_path)
            unique_files.append(f)
    
    unique_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    
    return unique_files


def main():
    print("\n╔══════════════════════════════════════════╗")
    print("║      Email Sender - Manual Mode          ║")
    print("╚══════════════════════════════════════════╝")
    
    file_path = None
    
    # Check if file path was provided
    if len(sys.argv) >= 2:
        # User provided a file path
        file_path = Path(sys.argv[1])
    else:
        # No file path provided - try to find most recent Excel file
        print("\n🔍 No file specified. Searching for Excel files...")
        
        try:
            available_files = list_available_excel_files()
            
            if not available_files:
                print("\n❌ No Excel files found!")
                print("\n📝 USAGE:")
                print("   python email_sender.py <file_path>")
                print("\n📋 EXAMPLES:")
                print("   python email_sender.py myfile.xlsx")
                print("   python email_sender.py output/coingecko_all_data_20240203_143022.xlsx")
                print('   python email_sender.py "path with spaces/report.xlsx"')
                sys.exit(1)
            
            if len(available_files) == 1:
                # Only one file found - use it automatically
                file_path = available_files[0]
                print(f"\n✅ Found 1 Excel file: {file_path.name}")
                print(f"   Location: {file_path.parent if file_path.parent != Path('.') else 'current directory'}")
                print(f"   Modified: {datetime.fromtimestamp(file_path.stat().st_mtime).strftime('%d %b %Y at %H:%M')}")
                
                # Ask for confirmation
                response = input("\n📧 Send this file? (Y/n): ").strip().lower()
                if response and response not in ['y', 'yes']:
                    print("\n❌ Cancelled by user.\n")
                    sys.exit(0)
            
            else:
                # Multiple files found - let user choose
                print(f"\n📋 Found {len(available_files)} Excel files:\n")
                
                for idx, f in enumerate(available_files, 1):
                    file_size_mb = f.stat().st_size / (1024 * 1024)
                    mod_time = datetime.fromtimestamp(f.stat().st_mtime).strftime('%d %b %Y, %H:%M')
                    location = f.parent if f.parent != Path('.') else 'current dir'
                    
                    print(f"   {idx}. {f.name}")
                    print(f"      └─ {file_size_mb:.2f} MB | {mod_time} | {location}")
                
                # Ask user to choose
                while True:
                    try:
                        choice = input(f"\n📧 Enter number to send (1-{len(available_files)}) or 'q' to quit: ").strip()
                        
                        if choice.lower() in ['q', 'quit', 'exit']:
                            print("\n❌ Cancelled by user.\n")
                            sys.exit(0)
                        
                        choice_num = int(choice)
                        if 1 <= choice_num <= len(available_files):
                            file_path = available_files[choice_num - 1]
                            break
                        else:
                            print(f"   ⚠️  Please enter a number between 1 and {len(available_files)}")
                    except ValueError:
                        print("   ⚠️  Please enter a valid number or 'q' to quit")
        
        except FileNotFoundError as e:
            print(f"\n❌ {e}")
            print("\n📝 USAGE:")
            print("   python email_sender.py <file_path>")
            print("\n📋 EXAMPLES:")
            print("   python email_sender.py myfile.xlsx")
            print("   python email_sender.py output/coingecko_all_data_20240203_143022.xlsx")
            sys.exit(1)
    
    # Optional: Custom recipient, subject, or body from command line
    recipient = sys.argv[2] if len(sys.argv) > 2 else None
    subject = sys.argv[3] if len(sys.argv) > 3 else None
    body = sys.argv[4] if len(sys.argv) > 4 else None
    
    # Send the email
    success = send_email_with_attachment(
        file_path=file_path,
        recipient=recipient,
        subject=subject,
        body=body
    )
    
    # Exit with appropriate code
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()