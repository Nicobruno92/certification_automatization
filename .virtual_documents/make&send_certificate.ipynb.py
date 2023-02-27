from utils import send_mail, eliminate_accents, certificate_maker

import pandas as pd


folder_path = "Certificates/Semana_psicologia_cientifica/"
list_path = folder_path + "lista_alumnos.xlsx"
certificate_folder = folder_path + "Certificates/"
certificate_template = folder_path + "certificate_template.jpeg"
font_file = folder_path + "D-DIN-Bold.ttf"  # an example of an special font


student_list = pd.read_excel(list_path)  # read excel files


email = "Correo electrónico"
first_name = "Nombre"
last_name = "Apellido"


#basic colors
white = '#FFFFFF'
black = '#000000'

# NT colors
grey = "#21201F"
green = "#9AC529"
lblue = "#42B9B2"
pink = "#DE237B"
orange = "#F38A31"

text_color = white

location_text = (160, 270)
font_size = 100


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


subject = "Test email automatization of certificates"
sender_email = "noreply.neurotransmitiendo@gmail.com"
password = input("Type your password and press enter: ")


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


for index, row in student_list.iterrows():

    # Information from the list
    receiver_email = row[email]

    full_name = row[first_name] + " " + row[last_name]

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
