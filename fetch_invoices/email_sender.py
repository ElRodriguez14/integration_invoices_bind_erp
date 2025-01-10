import smtplib
import re
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def get_client_email(client_name):
    # Diccionario con los correos electrónicos de los clientes
    client_emails = {
        "ABASTECEDORA DE CARNICOS DEL SURESTE": "germanrodriguez1@gmail.com",
        "Client2": "client2@example.com",
        # Agregar más clientes según sea necesario
    }
    return client_emails.get(client_name, None)  # Retorna None si no encuentra el cliente


def clean_filename(client_name):
    cleaned_name = re.sub(r'[^a-zA-Z0-9-_]', '_', client_name)
    return cleaned_name


def send_email_with_attachment(to_email, subject, body, attachment_path, smtp_server, smtp_port, smtp_user, smtp_password):
    # Crear el mensaje del correo
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = to_email
    msg['Subject'] = subject

    # Cuerpo del mensaje
    msg.attach(MIMEText(body, 'plain'))

    # Adjuntar el archivo CSV
    part = MIMEBase('application', "octet-stream")
    with open(attachment_path, "rb") as file:
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
    msg.attach(part)

    # Conectar al servidor SMTP y enviar el correo
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)
        text = msg.as_string()
        server.sendmail(smtp_user, to_email, text)
        server.quit()
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")


def send_invoices_to_clients(organized_invoices, smtp_server, smtp_port, smtp_user, smtp_password):
    for client, invoices_list in organized_invoices.items():
        file_name = f"{clean_filename(client)}_invoices.csv"
        file_path = os.path.join("output", file_name)

        # Suponiendo que en los datos de factura, el cliente tiene un campo de correo
        client_email = get_client_email(client)
        if client_email:
            subject = f"Invoices for {client}"
            body = f"Dear {client},\n\nPlease find attached the invoices for your recent transactions."
            send_email_with_attachment(client_email, subject, body, file_path, smtp_server, smtp_port, smtp_user, smtp_password)
        else:
            print(f"Email not found for client {client}")

