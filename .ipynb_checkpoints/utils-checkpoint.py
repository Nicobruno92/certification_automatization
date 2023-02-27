#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Nov 25 20:02:47 2020

@author: nicobruno
@email: nicobruno92@gmail.com
"""
#  Email
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.base import MIMEBase

import pandas as pd
import os

# For certificate
from PIL import Image, ImageDraw, ImageFont



def format_mail(sender_email, receiver_email, subject, body, filename):
    
    """Returns the mail with the format to be send"""
    
    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
#    message["To"] = receiver_email
    message["Subject"] = subject
    message["Bcc"] = receiver_email  # Recommended for mass emails
    
    # Turn these into plain/html MIMEText objects
    body_mail = MIMEText(body, "html")
    
    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(body_mail)
    
    if filename == False:
        None
        
    else:
        
        # Open PDF file in binary mode
        with open(filename, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        
        # Encode file in ASCII characters to send by email    
        encoders.encode_base64(part)
        
        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {os.path.basename(filename)}",
        )
        
        # Add attachment to message and convert message to string
        message.attach(part)
        
    
    return message.as_string()
    
def send_mail(sender_email, password,receiver_email, subject, body, attachment = False):
    
    """Sends the email
    sender_email: email of the sender
    receiver: email that will receive
    subject: subject of the email
    body: text content of the email in HTML format
    attachment: filename to be attached to the email
    """
    
    smtp_server = "smtp.gmail.com"
    port = 587  # For starttls
    
    # Message
    message = format_mail(sender_email, receiver_email, subject, body, filename = attachment)
    
    # Ask for passwords
    # password = input("Type your password and press enter: ")
    
    # Create a secure SSL context
    context = ssl.create_default_context()
    
    # Try to log in to server and send email
    try:
        server = smtplib.SMTP(smtp_server,port)
        server.ehlo() # Can be omitted
        server.starttls(context=context) # Secure the connection
        server.ehlo() # Can be omitted
        server.login(sender_email, password)
        # Send email here
        server.sendmail(sender_email, receiver_email, message)
    except Exception as e:
        # Print any error messages to stdout
        print(e)
    finally:
        server.quit() 
        
def eliminate_accents(name):
    
    name = name.lower()
    
    name = name.replace("á", "a")
    name = name.replace("é", "e")
    name = name.replace("í", "i")
    name = name.replace("ó","o")
    name = name.replace("ú", "u")
    name = name.replace("ü", "u")
    name = name.replace("ñ", "n")
    
    return name


def certificate_maker(certificate_template,student_name,text_color, location_text, font_name, text_size, save_path = ''):
    # Fuente y tamaño
    font = ImageFont.truetype(font_name, text_size)
    
    #open file
    im = Image.open(certificate_template)
    # crear imagen para el certificado
    d = ImageDraw.Draw(im)
    # composicion de la imagen del certidicado
    d.text(xy = location_text, text = student_name, fill=text_color, font=font, anchor = 'mm', align = 'center')
    
    #guardar el certificado con nombre y apellido
    im.save( save_path +'/certificado_' + eliminate_accents(student_name) + '.pdf')
        

def font_size_by_name(student_name, max_size):
    if len(student_name) > 30:
        size = max_size / 2
    elif len(student_name) > 20:
        size = max_size * 2/3
    else:
        size = max_size
        
    return int(size)