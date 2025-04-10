# Automatization of certification process for courses and conferences
Author: Nicolás Bruno

email: nicobruno92@gmail.com

repository: https://github.com/Nicobruno92/certification_automatization

## Overview
This repository contains scripts to automate the process of creating and sending certificates to participants of courses and conferences. The system handles:
- Creating personalized certificates with participant names
- Sending personalized emails with attached certificates
- Managing batches of certificates for different types of participants (attendees, presenters, etc.)

## Dependencies
Libraries to have installed:
- pillow - For image manipulation and certificate creation
- pandas - For handling data from spreadsheets
- email - For email composition and sending

The utility functions are in `utils.py` which must be in the same folder as the main scripts.

## Core Functionality (utils.py)

The `utils.py` file contains several key functions:

```python
from utils import send_mail, eliminate_accents, certificate_maker

import pandas as pd
```

### Key Functions in utils.py

1. **certificate_maker()** - Creates personalized certificates
   - Takes a certificate template image and adds a participant's name
   - Handles formatting, text positioning, and color
   - Saves the certificate as a PDF file

2. **send_mail()** - Sends emails with optional attachments
   - Handles SMTP connection and email composition
   - Supports HTML content for rich email formatting
   - Attaches certificates to emails

3. **eliminate_accents()** - Normalizes text by removing accents from characters
   - Used for filename consistency

4. **font_size_by_name()** - Adjusts font size based on name length
   - Ensures names fit properly on certificate templates

## Available Scripts and How to Use Them

### 1. Basic Certificate Workflow (make&send_certificate.ipynb)

This Jupyter notebook demonstrates the basic workflow for creating and sending certificates. It's an interactive way to understand the core functionality.

#### How to use:
1. Open the notebook in Jupyter
2. Execute each cell sequentially
3. Customize parameters as needed 
4. Test with a single certificate before sending to all participants

### 2. Creating Certificates 

#### A. For attendees (create_assistant_certificates.py)

This script creates certificates for event attendees/assistants.

**How to use:**
```bash
python create_assistant_certificates.py
```

**Configuration needed:**
- Modify the script to point to your attendee list (Excel/CSV)
- Set the certificate template path
- Configure the text position, font, and colors
- Define output directory for saving certificates

#### B. For presenters (create_exposition_certificates.py)

This script creates certificates for presenters at your event.

**How to use:**
```bash
python create_exposition_certificates.py
```

**Configuration needed:**
- Update the presenter list path
- Set the appropriate certificate template
- Adjust text formatting (may differ from attendee certificates)
- Define output directory

### 3. Sending Certificates

#### A. To attendees (send_certificates_asistentes.py)

This script sends certificates to event attendees via email.

**How to use:**
```bash
# Test SMTP connection only
python send_certificates_asistentes.py --test

# Test email to a specific address
python send_certificates_asistentes.py --test your.email@example.com

# Send certificates to all attendees
python send_certificates_asistentes.py
```

**Configuration required:**
- Create a `config_gmail.py` file with your email settings:
  ```python
  # config_gmail.py example
  EMAIL_USER = "your.email@gmail.com"
  EMAIL_PASSWORD = "your-app-password"  # Use app password for Gmail
  EMAIL_FROM = "Your Name <your.email@gmail.com>"
  REPLY_TO = "support@yourorganization.com"
  SMTP_SERVER = "smtp.gmail.com"
  SMTP_PORT = 587
  DOMAIN = "yourorganization.com"
  ```
- Create an `email_template_asistentes.py` file with your email content:
  ```python
  # email_template_asistentes.py example
  SUBJECT = "Your Certificate for [Event Name]"
  BODY = """
  <html>
  <body>
      <p>Hello {attendee_name},</p>
      <p>Thank you for attending our event. Your certificate is attached.</p>
      <p>Best regards,<br>Event Team</p>
  </body>
  </html>
  """
  ```
- Ensure certificates are generated and placed in the `certificates_dir` path specified in the script

#### B. To presenters (send_certificates_presentations.py)

This script sends certificates to event presenters.

**How to use:**
```bash
# Same options as the attendee version
python send_certificates_presentations.py [options]
```

**Configuration:**
- Similar to the attendee version but may use different email templates and certificate paths
- Create appropriate config and template files as described above

## Set path to folders
Set the path to the folder containing the necessary files such as:
- the certificate template (always in jpg)   
- the student list
- font file (if necessary)
- folder to save the certificates
- etc


```python
folder_path = "Automatization Example/"
list_path = folder_path + "lista_alumnos.xlsx"
certificate_folder = folder_path + "Certificates/"
certificate_template = folder_path + "certificate_template.jpg"
font_file = folder_path + "D-DIN-Bold.ttf"  # an example of an special font
```

## Read list of students
Set the name of the columns with student first name, last name and email


```python
student_list = pd.read_excel(list_path)  # read excel files
# sudent_list = pd.read_csv(list_path) #read csv files

email = "mail"
first_name = "Nombre"
last_name = "Apellido"
```

## Set Certificate parameters 
- Location of the name on the template
- Color of the text
- Size of the text
- Fonts
- etc

Finally makes a certificate to see if it's ok


```python
# NT colors
grey = "#21201F"
green = "#9AC529"
lblue = "#42B9B2"
pink = "#DE237B"
orange = "#F38A31"

text_color = grey

location_text = (810, 470)
font_size = 63
```


```python
# example certificate
certificate_maker(
    certificate_template=certificate_template,
    student_name="Juan Pérez",
    text_color=text_color,
    location_text=location_text,
    font_name=font_file,
    text_size=font_size,
    save_path=certificate_folder,
)
```

## Set email parameters

For this it is important to use a throwaway email. Because security parameters will complicate everything


```python
subject = "Test email automatization of certificates"
sender_email = "noreply.neurotransmitiendo@gmail.com"
password = input("Type your password and press enter: ")
```

## Email text

What it is in {} is what is going to be replaced afterwards by the name inside the loop function in order to personalize it.

The text is in HTML. 
- In order to end a paragraph add a break "\<br>"
- \<b> - <b> Bold text
- \<strong> <strong> - Important text
- \<i><i> - Italic text
- \<em><em> - Emphasized text
- \<mark><mark> - Marked text
- \<del><del> - Deleted text
- \<ins><ins> - Inserted text
- \<sub><sub> - Subscript text
- \<sup><sup> - Superscript text

## Setting Up Your Own Certificate System

To adapt this system for your own event:

1. **Prepare Your Data:**
   - Create an Excel/CSV file with participant information (name, email, etc.)
   - Design your certificate template as a JPG image
   - Choose appropriate fonts

2. **Configure Scripts:**
   - Update file paths to point to your data and template files
   - Adjust certificate text position and styling
   - Customize email templates

3. **Test Thoroughly:**
   - Generate sample certificates to check layout and formatting
   - Send test emails to yourself before mass sending
   - Verify that all certificates are correctly personalized

4. **Execute:**
   - Run certificate creation scripts first
   - Verify all certificates were created correctly
   - Run sending scripts with appropriate options

## Usage Example

To run the certificate creation and mailing process with the basic workflow:

1. Set up your data file (Excel or CSV) with participant information
2. Configure the certificate template and parameters
3. Test a single certificate generation to verify layout
4. Run the full process with:

```python
for index, row in student_list.iterrows():

    # Information from the list
    receiver_email = row[email]

    full_name = row[first_name] + " " + row[last_name]

    # Makes the certificate
    certificate_maker(
        certificate_template=certificate_template,
        student_name=full_name,
        text_color=text_color,
        location_text=location_text,
        font_name=font_file,
        text_size=font_size,
        save_path=certificate_folder,
    )

    # Prepares email
    body_formatted = body.format(name=row[first_name])

    attachment = (
        certificate_folder + "certificado_" + eliminate_accents(full_name) + ".pdf"
    )

    # send the email, good luck!
    send_mail(
        sender_email=sender_email,
        receiver_email=receiver_email,
        password=password,
        subject=subject,
        body=body_formatted,
        attachment=attachment,
    )
```

## Troubleshooting

- **Email Sending Failures**: Verify your SMTP settings and ensure you're using an app password for Gmail
- **Certificate Formatting Issues**: Adjust text position and font size based on your template
- **Filename Errors**: Check that the `eliminate_accents()` function correctly handles special characters in names
- **Missing Attachments**: Ensure certificates are being generated with correct names before sending

## WARNING

# <mark>BE CAREFUL! There is no going back!

Always double-check your email list and certificate templates before sending certificates to all participants. Once emails are sent, they cannot be recalled.
