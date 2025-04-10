import csv
import os
import smtplib
import sys
import time
import traceback
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import config_hostinger as config
import email_template


def normalize_name(name):
    """
    Normalize a name to match certificate filename format.
    This mimics the eliminate_accents function used in certificate creation.
    """
    if not isinstance(name, str):
        return name
        
    name = name.lower().strip()
    
    # Replace accents - copied from utils.py:eliminate_accents
    name = name.replace("á", "a")
    name = name.replace("é", "e")
    name = name.replace("í", "i")
    name = name.replace("ó", "o")
    name = name.replace("ú", "u")
    name = name.replace("ü", "u")
    name = name.replace("ñ", "n")
    
    # Strip spaces at the ends
    return name.strip()


def find_certificate(presenter_name, certificates_dir):
    """
    Find certificate PDF for a presenter.
    Uses the exact same normalization as the certificate creation script.
    """
    # Normalize the name exactly how it was normalized during certificate creation
    normalized_name = normalize_name(presenter_name)
    print(f"🔍 Looking for certificate for: '{presenter_name}'")
    print(f"   Normalized name: '{normalized_name}'")
    
    # First try direct filename construction
    expected_filename = f"certificado_{normalize_name(presenter_name)}.pdf"
    expected_path = os.path.join(certificates_dir, expected_filename)
    if os.path.exists(expected_path):
        print(f"✅ Found by direct filename: {expected_filename}")
        return expected_path
    
    # If direct construction fails, check all files
    for filename in os.listdir(certificates_dir):
        if filename.endswith('.pdf'):
            # Extract name from filename (remove prefix and extension)
            file_name = filename[len('certificado_'):-4].lower()
            file_normalized = normalize_name(file_name)
            
            # Try exact match
            if normalized_name == file_normalized:
                print(f"✅ Found exact match: {filename}")
                return os.path.join(certificates_dir, filename)
                
            # Try substring match
            if normalized_name in file_normalized or file_normalized in normalized_name:
                print(f"✅ Found substring match: {filename}")
                return os.path.join(certificates_dir, filename)
    
    # If we reach here, no certificate was found
    print(f"❌ No certificate found for '{presenter_name}'")
    return None


def test_smtp_connection():
    """Test SMTP connection and credentials."""
    print("\n--- Testing SMTP Connection ---")
    print(f"Server: {config.SMTP_SERVER}")
    print(f"Port: {config.SMTP_PORT}")
    print(f"User: {config.EMAIL_USER}")
    
    try:
        # Try to connect to the SMTP server
        with smtplib.SMTP_SSL(config.SMTP_SERVER, config.SMTP_PORT) as server:
            print("✅ Connected to SMTP server")
            
            # Try to login
            server.login(config.EMAIL_USER, config.EMAIL_PASSWORD)
            print("✅ Login successful")
            
            return True
    
    except Exception as e:
        print(f"❌ SMTP Connection Error: {str(e)}")
        traceback.print_exc()
        return False


def send_certificate(presenter_name, title, authors, email, certificate_path):
    """Send certificate to presenter via email."""
    try:
        print(f"\nPreparing email to: {email}")
        
        # Create message
        msg = MIMEMultipart()
        msg['From'] = config.EMAIL_FROM
        msg['To'] = email
        msg['Subject'] = email_template.SUBJECT
        if hasattr(config, 'REPLY_TO'):
            msg['Reply-To'] = config.REPLY_TO
        
        # Add additional headers to reduce spam probability
        msg['Message-ID'] = f"<{int(time.time())}@{config.EMAIL_FROM.split('@')[1]}>"
        msg['X-Mailer'] = "Python Email Sender"
        
        # Attach message body
        body = email_template.BODY.format(
            presenter_name=presenter_name,
            title=title,
            authors=authors
        )
        msg.attach(MIMEText(body, 'html'))
        
        # Attach certificate
        print(f"Attaching certificate: {os.path.basename(certificate_path)}")
        try:
            with open(certificate_path, 'rb') as file:
                attachment = MIMEApplication(file.read(), _subtype="pdf")
                attachment.add_header(
                    'Content-Disposition', 
                    'attachment',
                    filename=os.path.basename(certificate_path)
                )
                msg.attach(attachment)
                print("✅ Certificate attached successfully")
        except Exception as e:
            print(f"❌ Error attaching certificate: {str(e)}")
            raise
        
        # Connect to SMTP server and send
        print("Connecting to SMTP server...")
        with smtplib.SMTP_SSL(config.SMTP_SERVER, config.SMTP_PORT) as server:
            print("Logging in...")
            server.login(config.EMAIL_USER, config.EMAIL_PASSWORD)
            print("Sending message...")
            server.send_message(msg)
            print("Message sent successfully")
            
        print(f"✅ Certificate sent to {presenter_name} ({email})")
        return True
    
    except smtplib.SMTPRecipientsRefused as e:
        print(f"❌ Recipient refused: {email}")
        print(f"SMTP Error: {str(e)}")
        return False
    
    except smtplib.SMTPException as e:
        print(f"❌ SMTP Error sending to {presenter_name} ({email})")
        print(f"SMTP Error: {str(e)}")
        
        # Check if it's a sending limit error
        if "exceeded" in str(e).lower() or "limit" in str(e).lower():
            print("⚠️ Hostinger sending limit reached. Save progress and try again later.")
            raise e  # Re-raise to handle in the main function
        
        return False
    
    except Exception as e:
        print(f"❌ Error sending to {presenter_name} ({email})")
        print(f"Error: {str(e)}")
        traceback.print_exc()
        return False


def send_test_email(test_email):
    """Send a test email without certificate to verify SMTP works."""
    try:
        print(f"\n--- Sending test email to {test_email} ---")
        
        # Create message
        msg = MIMEMultipart()
        msg['From'] = config.EMAIL_FROM
        msg['To'] = test_email
        msg['Subject'] = "Test Email - Congreso Neurociencias"
        if hasattr(config, 'REPLY_TO'):
            msg['Reply-To'] = config.REPLY_TO
        
        # Add additional headers
        domain = config.EMAIL_FROM.split('@')[1]
        msg['Message-ID'] = f"<test-{int(time.time())}@{domain}>"
        msg['X-Mailer'] = "Python Email Sender"
        
        # Attach message body
        body = """
        <html>
        <body>
            <p>Este es un correo de prueba para verificar la configuración.</p>
            <p>Si recibiste este correo, la configuración es correcta.</p>
        </body>
        </html>
        """
        msg.attach(MIMEText(body, 'html'))
        
        # Connect to SMTP server and send
        with smtplib.SMTP_SSL(config.SMTP_SERVER, config.SMTP_PORT) as server:
            server.login(config.EMAIL_USER, config.EMAIL_PASSWORD)
            server.send_message(msg)
            
        print(f"✅ Test email sent to {test_email}")
        return True
    
    except Exception as e:
        print(f"❌ Error sending test email: {str(e)}")
        traceback.print_exc()
        return False


def load_sent_certificates():
    """Load the list of already sent certificates from JSON file."""
    sent_file = "congreso_neurociencias/sent_certificates_hostinger.json"
    if os.path.exists(sent_file):
        try:
            with open(sent_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"❌ Error loading sent certificates: {str(e)}")
            return []
    else:
        return []


def save_sent_certificates(sent_list):
    """Save the list of sent certificates to JSON file."""
    sent_file = "congreso_neurociencias/sent_certificates_hostinger.json"
    try:
        with open(sent_file, 'w', encoding='utf-8') as f:
            json.dump(sent_list, f, ensure_ascii=False, indent=2)
        print(f"✅ Saved progress to {sent_file}")
    except Exception as e:
        print(f"❌ Error saving sent certificates: {str(e)}")


def is_certificate_sent(presenter_name, title, sent_certificates):
    """Check if the certificate has already been sent using flexible matching."""
    normalized_presenter = normalize_name(presenter_name)
    normalized_title = normalize_name(title)
    
    # Create candidate IDs with different formats to try
    candidate_ids = [
        f"{normalized_presenter}:{normalized_title}",
        # Alternative formats that might be in the JSON
        normalized_presenter
    ]
    
    # Try exact matching first
    for cert_id in sent_certificates:
        cert_id_norm = normalize_name(cert_id)
        for candidate in candidate_ids:
            if cert_id_norm == candidate:
                return True
    
    # Try partial matching if exact match failed
    for cert_id in sent_certificates:
        cert_id_norm = normalize_name(cert_id)
        
        # Split the ID to compare name and title separately
        if ':' in cert_id_norm:
            cert_name, cert_title = cert_id_norm.split(':', 1)
            
            # Check if both name and at least part of the title match
            name_match = (normalized_presenter in cert_name or 
                         cert_name in normalized_presenter)
            title_match = (normalized_title in cert_title or 
                          cert_title in normalized_title)
            if name_match and title_match:
                return True
    
    return False


def main():
    # Parse command-line arguments
    if len(sys.argv) > 1 and sys.argv[1] == "--test":
        if len(sys.argv) > 2:
            # Test email to a specific address
            test_smtp_connection()
            send_test_email(sys.argv[2])
        else:
            # Just test SMTP connection
            test_smtp_connection()
        return
    
    # Normal operation - send certificates
    certificates_dir = "congreso_neurociencias/certificados_expositores"
    csv_path = "congreso_neurociencias/Presentadores Congreso.csv"
    
    # First test SMTP connection
    if not test_smtp_connection():
        print("❌ SMTP connection failed. Check your credentials and settings.")
        return
    
    # Load previously sent certificates
    sent_certificates = load_sent_certificates()
    print(f"ℹ️ Found {len(sent_certificates)} previously sent certificates")
    
    # Debug: Print all loaded certificates to verify
    print("\n--- Already sent certificates ---")
    for cert in sent_certificates:
        print(f"✓ {cert}")
    print("--- End of sent certificates list ---\n")
    
    # Track results
    success_count = 0
    error_count = 0
    not_found_count = 0
    skipped_count = 0
    not_found_presenters = []
    
    # Process CSV file
    print("\n--- Processing CSV file ---")
    
    try:
        # Try opening the file as regular CSV
        with open(csv_path, 'r', encoding='utf-8') as csvfile:
            # First try reading as CSV with different delimiters
            delimiter = None
            for test_delimiter in [',', ';']:
                try:
                    print(f"Trying delimiter: '{test_delimiter}'")
                    csvfile.seek(0)  # Reset file position
                    reader = csv.reader(csvfile, delimiter=test_delimiter)
                    header = next(reader)  # Skip header
                    print(f"Header has {len(header)} columns: {header}")
                    
                    # Read first row to test
                    first_row = next(reader)
                    print(f"First row has {len(first_row)} columns: {first_row}")
                    
                    if len(first_row) >= 4:
                        print(f"✅ Found valid delimiter: '{test_delimiter}'")
                        delimiter = test_delimiter
                        break
                except Exception as e:
                    print(f"Error with delimiter '{test_delimiter}': {str(e)}")
                    continue
            
            if not delimiter:
                print("❌ Could not determine CSV delimiter. Defaulting to ','")
                delimiter = ','
            
            # Reset file and read with the correct delimiter
            csvfile.seek(0)
            reader = csv.reader(csvfile, delimiter=delimiter)
            next(reader)  # Skip header
            
            # Process rows
            for row in reader:
                try:
                    # Minimum required fields: title, presenter_name, authors, email
                    if len(row) >= 4:
                        title = row[0]
                        presenter_name = row[1]
                        authors = row[2]
                        email = row[3]
                    else:
                        # Try to read raw line and parse manually
                        print(f"Row too short, trying manual parsing: {row}")
                        continue
                    
                    if not email.strip():
                        print(f"⚠️ No email found for {presenter_name}. Skipping.")
                        continue
                    
                    # Cleanup any quotes in fields
                    title = title.strip('"').strip()
                    presenter_name = presenter_name.strip('"').strip()
                    authors = authors.strip('"').strip()
                    email = email.strip('"').strip()
                    
                    # Create unique ID for this certificate
                    unique_id = f"{normalize_name(presenter_name)}:{normalize_name(title)}"
                    
                    # Use the more reliable matching function
                    print(f"Checking if already sent: {presenter_name} - {title}")
                    if is_certificate_sent(presenter_name, title, sent_certificates):
                        print(f"ℹ️ Already sent to {presenter_name} for '{title}'. Skipping.")
                        skipped_count += 1
                        continue
                    
                    print(f"Processing: {presenter_name} <{email}>")
                    
                    # Find certificate
                    cert_path = find_certificate(presenter_name, certificates_dir)
                    
                    if cert_path:
                        # Send certificate
                        try:
                            if send_certificate(
                                presenter_name, title, authors, email, cert_path
                            ):
                                success_count += 1
                                # Add to sent certificates list
                                sent_certificates.append(unique_id)
                                save_sent_certificates(sent_certificates)
                                
                                # Add delay between emails to avoid spam filters
                                if success_count % 10 == 0:  # Every 10 emails
                                    print("Waiting 60 seconds to avoid limits...")
                                    time.sleep(60)
                                else:
                                    print("Waiting 20 seconds between emails...")
                                    time.sleep(10)  # Longer pause between emails
                            else:
                                error_count += 1
                        except smtplib.SMTPException as e:
                            if "exceeded" in str(e).lower() or "limit" in str(e).lower():
                                print("\n⚠️ Hostinger sending limit reached")
                                print("Progress saved. Run the script again later to continue.")
                                break
                            else:
                                error_count += 1
                    else:
                        print(f"⚠️ Certificate not found for {presenter_name}")
                        not_found_count += 1
                        not_found_presenters.append(presenter_name)
                except Exception as e:
                    print(f"❌ Error processing row: {str(e)}")
                    traceback.print_exc()
                    error_count += 1
    
    except Exception as e:
        print(f"❌ Error processing CSV file: {str(e)}")
        traceback.print_exc()
        return
    
    # Print summary
    print("\n--- Summary ---")
    print(f"Certificates sent: {success_count}")
    print(f"Skipped (already sent): {skipped_count}")
    print(f"Errors: {error_count}")
    print(f"Certificates not found: {not_found_count}")
    print(f"Total sent to date: {len(sent_certificates)}")
    
    if not_found_presenters:
        print("\nPresenters without certificates:")
        for presenter in not_found_presenters:
            print(f"- {presenter}")


if __name__ == "__main__":
    main() 