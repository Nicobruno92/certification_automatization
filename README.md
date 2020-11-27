# Automatization of certifaction process for courses and conferences
Author: Nicolás Bruno

email: nicobruno92@gmail.com

repository: https://github.com/Nicobruno92/certification_automatization

## Dependencies
Library to have installed:
pillow, pandas, email

file utils.py in same folder than this script. 


```python
from utils import send_mail, eliminate_accents, certificate_maker

import pandas as pd
```

## Set path to folders
Set the path to the folder containing the necessary files sucha as
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

For this it is important to use a throaway email. Because security parameters will complicate everything


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


```python
body = """\
            <html>
              <body>
                <p>
                    Hi {name},<br>
                    How are you doing?<br>
                    This email that you are receiving is a test email to try the automatization of 
                    creating and sending certificates. <br>

                    You can visit our website <a href="http://www.neurotransmitiendo.org">here</a>
                </p>
              </body>
            </html>
            """
```

## Send a test email


```python
receiver_email = "noreply.neurotransmitiendo@gmail.com"
attachment = certificate_folder + "certificado_" + "Juan Perez" + ".pdf"

send_mail(
    sender_email=sender_email,
    receiver_email=receiver_email,
    password=password,
    subject=subject,
    body=body,
    attachment=attachment,
)
```

## Create and send the certificate to all students in the list.

# <mark>BE CAREFUL! There is no going back!


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

    # PRepares email
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
