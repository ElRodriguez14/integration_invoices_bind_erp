import smtplib
import re
import os
import pandas as pd

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def get_client_email(client_name):
    # Diccionario con los correos electrónicos de los clientes
    client_emails = {
        "CARNES Y ABARROTES A A A": "germanrodriguez1@gmail.com",
        "Client2": "client2@example.com",
        # Agregar más clientes según sea necesario
    }
    return client_emails.get(client_name, None)  # Retorna None si no encuentra el cliente


def clean_filename(client_name):
    cleaned_name = re.sub(r'[^a-zA-Z0-9-_]', '_', client_name)
    return cleaned_name


def send_email_with_attachment(to_email, subject, html_body, attachment_paths, smtp_server, smtp_port, smtp_user, smtp_password):
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = to_email
    msg['Subject'] = subject

    # Cuerpo del mensaje en HTML
    msg.attach(MIMEText(html_body, 'html'))

    # Adjuntar cada archivo CSV
    for attachment_path in attachment_paths:
        part = MIMEBase('application', "octet-stream")
        with open(attachment_path, "rb") as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
        msg.attach(part)

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, to_email, msg.as_string())
        server.quit()
        print(f"Email sent successfully to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")


def csv_to_html_table(csv_path, moneda):
    # Leer el archivo CSV con pandas
    df = pd.read_csv(csv_path)
    df = df.fillna("")
    # Convertir el DataFrame en tabla HTML
    html_table = df.to_html(index=False, border=1, classes="table", justify="center")
    # Agregar un título a la tabla según la moneda
    title = f"<h3>Facturas en {'Pesos Mexicanos' if moneda == 'MXN' else 'Dólares Americanos'}</h3>"
    return title + html_table


def generate_file_name(client_name, currency):
    """
    Genera un nombre de archivo basado en el nombre del cliente y la moneda.

    Args:
        client_name (str): Nombre del cliente.
        currency (str): Moneda (USD o MXN).

    Returns:
        str: Nombre del archivo generado.
    """
    # Reemplazar espacios por "_"
    sanitized_client_name = client_name.replace(" ", "_")

    # Construir el nombre del archivo con moneda y sufijo
    file_name = f"{sanitized_client_name}_{currency}_invoices.csv"

    return file_name


def send_invoices_to_clients(organized_invoices, smtp_server, smtp_port, smtp_user, smtp_password):
    for client, invoices_list in organized_invoices.items():
        client_email = get_client_email(client)
        if client_email:
            # Generar cuerpo del correo en HTML
            body_html = f"<p>Estimado {client},</p><p>Por favor, encuentre las facturas adjuntas en este correo:</p>"
            attachment_paths = []

            for file_name in invoices_list:
                # Determinar moneda a partir del nombre del archivo
                moneda = "MXN" if "MXN" in file_name else "USD" if "USD" in file_name else "Desconocida"
                file_name_client = generate_file_name(client, file_name)

                file_path = os.path.join("output", file_name_client)
                # Generar tabla HTML y agregar al cuerpo del correo
                body_html += csv_to_html_table(file_path, moneda)
                # Agregar archivo a la lista de adjuntos
                attachment_paths.append(file_path)

            # Enviar correo con tablas y archivos adjuntos
            subject = f"Facturas para {client}"
            send_email_with_attachment(client_email, subject, body_html, attachment_paths, smtp_server, smtp_port,
                                       smtp_user, smtp_password)
        else:
            print(f"Email not found for client {client}")

