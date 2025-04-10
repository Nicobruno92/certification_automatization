import csv
import os
import smtplib
import sys
import time
import traceback
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import config_gmail as config
import email_template_asistentes as email_template


def normalize_name(name):
    """Normalize a name to match certificate filename format."""
    return name.lower().strip()


def find_certificate(attendee_name, certificates_dir):
    """Find certificate PDF for an attendee."""
    normalized_name = normalize_name(attendee_name)
    
    for filename in os.listdir(certificates_dir):
        if filename.endswith('.pdf'):
            # Extract name from filename (remove prefix and extension)
            file_name = filename[len('certificado_'):-4].lower()
            if normalized_name in file_name or file_name in normalized_name:
                return os.path.join(certificates_dir, filename)
    
    return None


def test_smtp_connection():
    """Test SMTP connection and credentials."""
    print("\n--- Testing SMTP Connection ---")
    print(f"Server: {config.SMTP_SERVER}")
    print(f"Port: {config.SMTP_PORT}")
    print(f"User: {config.EMAIL_USER}")
    
    try:
        # Try to connect to the SMTP server
        with smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT) as server:
            print("✅ Connected to SMTP server")
            
            # Start TLS encryption
            server.starttls()
            print("✅ TLS encryption started")
            
            # Try to login
            server.login(config.EMAIL_USER, config.EMAIL_PASSWORD)
            print("✅ Login successful")
            
            return True
    
    except Exception as e:
        print(f"❌ SMTP Connection Error: {str(e)}")
        traceback.print_exc()
        return False


def send_certificate(attendee_name, email, certificate_path):
    """Send certificate to attendee via email."""
    try:
        print(f"\nPreparing email to: {email}")
        
        # Create message
        msg = MIMEMultipart()
        msg['From'] = config.EMAIL_FROM
        msg['To'] = email
        msg['Subject'] = email_template.SUBJECT
        msg['Reply-To'] = config.REPLY_TO
        
        # Add additional headers to reduce spam probability
        msg['Message-ID'] = f"<{int(time.time())}@{config.DOMAIN}>"
        msg['X-Mailer'] = "Python Email Sender"
        
        # Attach message body
        body = email_template.BODY.format(
            attendee_name=attendee_name
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
        with smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT) as server:
            print("Starting TLS encryption...")
            server.starttls()
            print("Logging in...")
            server.login(config.EMAIL_USER, config.EMAIL_PASSWORD)
            print("Sending message...")
            server.send_message(msg)
            print("Message sent successfully")
            
        print(f"✅ Certificate sent to {attendee_name} ({email})")
        return True
    
    except smtplib.SMTPRecipientsRefused as e:
        print(f"❌ Recipient refused: {email}")
        print(f"SMTP Error: {str(e)}")
        return False
    
    except smtplib.SMTPException as e:
        print(f"❌ SMTP Error sending to {attendee_name} ({email})")
        print(f"SMTP Error: {str(e)}")
        return False
    
    except Exception as e:
        print(f"❌ Error sending to {attendee_name} ({email})")
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
        msg['Reply-To'] = config.REPLY_TO
        
        # Add additional headers
        msg['Message-ID'] = f"<test-{int(time.time())}@{config.DOMAIN}>"
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
        with smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT) as server:
            server.starttls()
            server.login(config.EMAIL_USER, config.EMAIL_PASSWORD)
            server.send_message(msg)
            
        print(f"✅ Test email sent to {test_email}")
        return True
    
    except Exception as e:
        print(f"❌ Error sending test email: {str(e)}")
        traceback.print_exc()
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
    certificates_dir = "congreso_neurociencias/certificados_asistentes"
    csv_filename = ("Inscripción al Primer Congreso Latinoamericano de "
                   "Neurociencias Cognitivas  (respuestas) - "
                   "Respuestas de formulario 1.csv")
    csv_path = f"congreso_neurociencias/{csv_filename}"
    
    # First test SMTP connection
    if not test_smtp_connection():
        print("❌ SMTP connection failed. Check your credentials and settings.")
        return
    
    # Track results
    success_count = 0
    error_count = 0
    not_found_count = 0
    not_found_attendees = []
    
    # Process CSV file
    with open(csv_path, 'r', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile, delimiter=';')
        next(reader)  # Skip header
        
        for row in reader:
            if len(row) >= 4:  # Ensure row has enough columns
                # Format: timestamp, email, country, name_and_surname
                _, email, _, attendee_name = row[:4]
                
                if not email.strip():
                    print(f"⚠️ No email found for {attendee_name}. Skipping.")
                    continue
                
                # Find certificate
                cert_path = find_certificate(attendee_name, certificates_dir)
                
                if cert_path:
                    # Send certificate
                    if send_certificate(attendee_name, email, cert_path):
                        success_count += 1
                        # Add delay between emails to avoid spam filters
                        if success_count % 5 == 0:  # Every 5 emails
                            print("Waiting 30 seconds to avoid limits...")
                            time.sleep(30)
                        else:
                            time.sleep(2)  # Short pause between emails
                    else:
                        error_count += 1
                else:
                    print(f"⚠️ Certificate not found for {attendee_name}")
                    not_found_count += 1
                    not_found_attendees.append(attendee_name)
    
    # Print summary
    print("\n--- Summary ---")
    print(f"Certificates sent: {success_count}")
    print(f"Errors: {error_count}")
    print(f"Certificates not found: {not_found_count}")
    
    if not_found_attendees:
        print("\nAttendees without certificates:")
        for attendee in not_found_attendees:
            print(f"- {attendee}")


if __name__ == "__main__":
    main() 