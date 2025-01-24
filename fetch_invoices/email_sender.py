import smtplib
import re
import os
import pandas as pd

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from email.mime.image import MIMEImage

import win32com.client as win32


def get_client_fiscal(client_name):
    # Diccionario con los correos electrónicos de los clientes
    client_fiscal = {
        "VENTA AL PUBLICO EN GENERAL - SP": "SAUL PEREZ",
        "VENTA AL PUBLICO EN GENERAL - RB": "RAUL BOVIO GUERRERO",
        "VENTA AL PUBLICO EN GENERAL  - CF": "CARNICOS LA FORTUNA",
        "VENTA AL PUBLICO EN GENERAL - IA": "ISA ALIMENTOS",
        "VENTA AL PUBLICO EN GENERAL - RS": "ROGELIO SERRANO",
        "VENTA AL PUBLICO EN GENERAL -AM": "ADRIAN MONTIEL",
        "VENTA AL PUBLICO EN GENERAL - AB": "ARTURO BARRADAS",
        "VENTA AL PUBLICO EN GENERAL  - RC": "RAUL COSME",
        "VENTAS AL PUBLICO EN GENERAL - DG": "DON GATO",
        "VENTA AL PUBLICO EN GENERAL - IE": "IVAN ESTRADA ALVARADO",
        "VENTA AL PUBLICO EN GENERAL  - MC": "MANUEL JESUS CORONADO SOSA",
        "VENTA AL PUBLICO EN GENERAL - FL": "FRANCISCO LOPEZ",
        "VENTA AL PUBLICO EN GENERAL - ML": "MARIO LOPEZ",
        "VENTA AL PUBLICO EN GENERAL - COP": "COPOCAR",
        "VENTA AL PUBLICO EN GENERAL - YC": "YESSICA CAICERO MURRIETA",
        "VENTA AL PUBLICO EN GENERAL - MN": "MIGUEL ANGEL NAVARRO DIAZ",
        "VENTA AL PUBLICO EN GENERAL - AQ": "AQUA",
        "VENTA AL PUBLICO EN GENERAL - COV": "COVAGO",
        "VENTA AL PUBLICO EN GENERAL - JR": "JASSO CRUZ",
        "VENTA AL PUBLICO EN GENERAL - GT": "GILDARDO TORRES",
        "VENTA AL PUBLICO EN GENERAL - MS": "MIGUEL ANGEL SANTIAGO HERNANDEZ",
        "VENTA AL PUBLICO EN GENERAL - JM": "JONATHAN MARTINEZ BUSTOS",
        "VENTA AL PUBLICO EN GENERAL - JC": "JAVIER CASTILLO",
        "VENTA AL PUBLICO EN GENERAL - JS": "JUAN SANTIAGO SANTIAGO",
        "VENTA AL PUBLICO EN GENERAL - OML": "OBRADOR Y EMPACADORA LARA",
        "VENTA AL PUBLICO EN GENERAL - FCH": "FERNANDO CHAVEZ",
        "VENTA AL PUBLICO GENERAL - OL": "OBRADOR Y EMPACADORA LARA",
        "VENTA AL PUBLICO EN GENERAL - EG": "ROSARIO GONZALEZ SERRANO",
        "VENTA AL PUBLICO EN GENERAL - JE": "JORGE ESQUIVEL CASTRO",
        "VENTA AL PUBLICO EN GENERAL - LA": "LUIS ARREGUIN",
        "VENTA AL PUBLICO EN GENERAL - AG": "ABELARDO GONZALEZ",
        "VENTA AL PUBLICO EN GENERAL - AL": "ALICIA LARIOS",
        "VENTA AL PUBLICO EN GENERAL - CO": "COPOCAR",
        "VENTA AL PUBLICO EN GENERAL - SHP": "SERGIO HERNANDEZ PONCE",
        "VENTA AL PUBLICO EN GENERAL - MG": "MARBUSTELL GRUPO COMERCIAL",
        "VENTA AL PUBLICO EN GENERAL - MV": "MIGUEL VALVERDE",
        "VENTA AL PUBLICO EN GENERAL  - MA": "JULIA MUÑOZ",
        "VENTA AL PUBLICO EN GENERAL - VE": "VICENTE ESTRADA",
        "VENTA AL PUBLICO EN GENERAL - JV": "JOSE ALFREDO VELA",
        "VENTA AL PUBLICO EN GENERAL - FC": "FRANCISCO COBOS"
    }
    return client_fiscal.get(client_name, [])


def get_client_emails(client_name):
    # Diccionario con los correos electrónicos de los clientes
    client_emails = {
        "VENTA AL PUBLICO EN GENERAL - RB": ["oscarduvan20667@gmail.com"],
        "CARNES Y ABARROTES A A A": ["oscarduvan20667@gmail.com"],

    }
    return client_emails.get(client_name, [])  # Retorna None si no encuentra el cliente


def clean_filename(client_name):
    cleaned_name = re.sub(r'[^a-zA-Z0-9-_]', '_', client_name)
    return cleaned_name


def send_email_with_attachment(to_email, subject, html_body, attachment_paths, smtp_server, smtp_port, smtp_user,
                               smtp_password, client_name, image_paths=None):

    """
    Code for GMAIL
        msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = to_email
    msg['Subject'] = subject

    # Cuerpo del mensaje en HTML
    msg.attach(MIMEText(html_body, 'html'))

    # Adjuntar cada archivo CSV

    #for attachment_path in attachment_paths:
    #    part = MIMEBase('application', "octet-stream")
    #    with open(attachment_path, "rb") as file:
    #        part.set_payload(file.read())
    #    encoders.encode_base64(part)
    #    part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
    #    msg.attach(part)

    try:
        # Agregar imágenes si las rutas son válidas
        if image_paths:
            for i, image_path in enumerate(image_paths):
                if os.path.exists(image_path):
                    with open(image_path, 'rb') as img_file:
                        img = MIMEImage(img_file.read())
                        # Agregar un Content-ID único para cada imagen
                        content_id = f"image_{i+1}"
                        img.add_header('Content-ID', f'<{content_id}>')
                        msg.attach(img)
                        print(f"Image {i+1} attached with CID: {content_id}")
                else:
                    print(f"Image {image_path} not found.")

        # Enviar correo
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, to_email, msg.as_string())
        server.quit()
        print(f"Email sent successfully for Client {client_name} to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")
    """

    # Code for Outlook

    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)


    # Seleccionar la cuenta de envío (si se especifica)
    # Cuenta a usar de las que tenga en Outlook registradas
    from_account = "oscar_rodriguez_1402@hotmail.com"

    if from_account:
        # Obtener todas las cuentas configuradas en Outlook
        accounts = outlook.Session.Accounts
        for account in accounts:
            if account.DisplayName == from_account:
                mail.SendUsingAccount = account
                break
        else:
            print(f"Cuenta '{to_email}' no encontrada, utilizando la cuenta predeterminada.")

    mail.to = to_email
    mail.Subject = subject
    mail.HTMLBody = html_body

    # Cuerpo del mensaje en HTML

    # Adjuntar cada archivo CSV
    #for attachment_path in attachment_paths:
    #    part = MIMEBase('application', "octet-stream")
    #    with open(attachment_path, "rb") as file:
    #        part.set_payload(file.read())
    #    encoders.encode_base64(part)
    #    part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
    #    msg.attach(part)

    try:
        # Agregar imágenes si las rutas son válidas
        for i, image_path in enumerate(image_paths):
            absolute_path = os.path.abspath(image_path)

            if os.path.exists(absolute_path):
                attachment = mail.Attachments.Add(absolute_path)
                # Asignar un Content-ID único
                content_id = f"image_{i + 1}"
                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E",
                                                        content_id)
                # Incluir la imagen en el HTML con su CID
                html_body += f'<img src="cid:{content_id}" style="display: block; margin: 10px auto;">'
                print(f"Image {i + 1} attached with CID: {content_id}")
            else:
                print(f"Image {image_path} not found.")

        mail.Send()

        print(f"Email sent successfully for Client {client_name} to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")


def csv_to_html_table(csv_path, name, moneda):
    # Leer el archivo CSV con pandas
    df = pd.read_csv(csv_path)
    df = df.fillna("")  # Llenar valores NaN con cadena vacía

    # Convertir columnas específicas
    columns_to_convert = ['Factura', 'PO', 'Dias Vencidos']
    for column in columns_to_convert:
        if column in df.columns:  # Verificar si la columna existe en el DataFrame
            df[column] = pd.to_numeric(df[column], errors='coerce').apply(
                lambda x: int(x) if isinstance(x, float) and x.is_integer() else "" if pd.isna(x) else x)

    # Formatear las columnas con decimales
    columns_with_decimals = ['Total', 'Balance']
    for column in columns_with_decimals:
        if column in df.columns:
            df[column] = df[column].apply(
                lambda x: (
                    f"({abs(float(x.strip('()'))):,.2f})" if isinstance(x, str) and x.startswith("(") and x.endswith(
                        ")")
                    else f"{float(x):,.2f}"  # Formato estándar de número con coma y punto
                ) if isinstance(x, (int, float, str)) and pd.notna(x) else x
            )

    # Convertir el DataFrame en tabla HTML con clases
    html_table = df.to_html(index=False, border=1, classes="table", justify="center")

    # Estilizar la primera fila (encabezados) y las columnas

    html_table = html_table.replace('<table ',
                                    '<table style="border-collapse: collapse; border: 2px solid black;" ')  # Borde grueso
    html_table = html_table.replace('<thead>',
                                    '<thead style="background-color: #307BDA; color: black; font-weight: bold;">')  # Color negro en el encabezado

    # Establecer un ancho fijo para todas las columnas (120px)
    html_table = html_table.replace('<th>', '<th style="width: 120px; text-align: center;">')
    html_table = html_table.replace('<td>', '<td style="width: 120px; text-align: center;">')


    # Título con la moneda
    title = f"<h3>Facturas de {name} en {'Pesos Mexicanos' if moneda == 'MXN' else 'Dólares Americanos'}</h3>"

    # Devuelvo el HTML final con el título y la tabla estilizada
    html_final = title + html_table
    return html_final


def generate_file_name(client_name, currency):
    # Reemplazar espacios por "_"
    sanitized_client_name = client_name.replace(" ", "_").replace("-", "_")

    # Construir el nombre del archivo con moneda y sufijo
    file_name = f"{sanitized_client_name}_{currency}_invoices.csv"

    return file_name


def send_invoices_to_clients(organized_invoices, smtp_server, smtp_port, smtp_user, smtp_password, list_clients):
    for client, invoices_list in organized_invoices.items():
        client_emails = get_client_emails(client)
        if client_emails:
            for email in client_emails:

                # Leer cada archivo CSV para calcular los balances vencidos
                attachment_paths = []

                #nuevo
                names = []
                names.append(client)
                client_variant = get_client_fiscal(client)
                if client_variant and client_variant in list_clients:
                    names.append(client_variant)
                else:
                    pass

                body_html = f"""
                <html>
                    <body>
                        <img src="cid:image_1" alt="Logo" style="width:150px; margin-top: 20px;">
                        <h3 style="text-align: left;">ESTADO DE CUENTA</h3>
                        <p>Estimado <strong>{client}</strong>,</p>
                        <p>Adjunto encontrará su estado de cuenta al día de hoy. Nuestro sistema muestra el siguiente balance vencido:</p>

                    </body>
                </html>
                """
                currency_types = ["MXN", "USD"]
                for currency in currency_types:
                    # Determinar moneda a partir del nombre del archivo

                    for name in names:

                        balance_vencido_mxn = 0
                        balance_vencido_usd = 0

                        file_name_client = generate_file_name(name, currency)
                        file_path = os.path.join("output", file_name_client)

                        if not os.path.exists(file_path):
                            print(f"Archivo {file_path} no encontrado. Continuando con la siguiente iteración.")
                            continue

                        # Leer el CSV para obtener el balance vencido
                        try:
                            df = pd.read_csv(file_path)
                            if "Balance" in df.columns and not df.empty:
                                last_balance = df["Balance"].iloc[-1]  # Último valor de la columna "Balance"
                                last_balance = float(last_balance) if pd.notna(last_balance) else 0

                                if currency == "MXN":
                                    balance_vencido_mxn += last_balance
                                elif currency == "USD":
                                    balance_vencido_usd += last_balance
                        except Exception as e:
                            print(f"Error reading {file_path}: {e}")
                            continue

                        # Generar tabla HTML para el archivo y agregar al cuerpo del correo
                        attachment_paths.append(file_path)

                        if balance_vencido_mxn != 0:
                            body_html += f"""
                                    <!-- Tabla para MXN -->
                                    <p>Balance Vencido para <strong>{name}</strong></p>
                                    <table style="border-collapse: collapse; text-align: left; margin-bottom: 15px;">
                                        <tr>
                                            <td style="padding: 4px; border: 2px solid black; background-color: #307BDA; color: #000000; width: 80px; height: 25px; text-align: center; white-space: nowrap;"><strong>MXN:</strong></td>
                                            <td style="padding: 4px; border: 2px solid black; width: 80px; height: 25px; text-align: center; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{balance_vencido_mxn:,.2f}</td>
                                        </tr>
                                    </table>


                            """
                        if balance_vencido_usd != 0:
                            body_html += f"""                    
                                    <!-- Tabla para USD -->
                                    <p>Balance Vencido para <strong>{name}</strong></p>
                                    <table style="border-collapse: collapse; text-align: left; margin-bottom: 15px;">
                                        <tr>
                                            <td style="padding: 4px; border: 2px solid black; background-color: #307BDA; color: #000000; width: 80px; height: 25px; text-align: center; white-space: nowrap;"><strong>USD:</strong></td>
                                            <td style="padding: 4px; border: 2px solid black; width: 80px; height: 25px; text-align: center; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{balance_vencido_usd:,.2f}</td>
                                        </tr>
                                    </table>
                            """

                today = datetime.now()
                format_date = today.strftime("%d/%m/%Y")
                subject = f"Estado de Cuenta para {client} - {format_date}"


                body_html += """
                    <p>Si el pago ha sido realizado, favor de omitir este mensaje.</p>
                    <p>Si usted tiene alguna pregunta sobre su estado de cuenta, por favor contactarse con nosotros.</p>
                    <p>Agradeciendo la atención a la presente, por su apoyo y continuo negocio.</p>

                """

                # Ruta de la imagen
                image_paths = ["images/first_logo.png", "images/second_logo.png"]

                attachment_paths = []

                for currency in currency_types:
                    for name in names:
                        file_name_client = generate_file_name(name, currency)

                        file_path = os.path.join("output", file_name_client)

                        if not os.path.exists(file_path):
                            print(f"Archivo {file_path} no encontrado. Continuando con la siguiente iteración.")
                            continue

                        # Generar tabla HTML y agregar al cuerpo del correo
                        body_html += csv_to_html_table(file_path, name, currency)
                        # Agregar archivo a la lista de adjuntos
                        attachment_paths.append(file_path)

                body_html += """
                    <p><br /><br /><strong>Cordial saludo</strong>, muchas gracias.</p>
                """
                body_html += f"""
                    <img src="cid:image_2" alt="Second Image" style="width:676px; margin-top: 20px; margin-bottom: 0;">
                    </body>
                    </html>
                    """

                # Enviar correo con tablas y archivos adjuntos
                send_email_with_attachment(email, subject, body_html, attachment_paths, smtp_server, smtp_port,
                                           smtp_user, smtp_password, client, image_paths)
        else:
            print(f"Email not found for client {client}")
