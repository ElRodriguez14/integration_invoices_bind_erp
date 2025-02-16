import smtplib
import re
import os
import pandas as pd
import imghdr

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from email.mime.image import MIMEImage



def get_client_emails(client_name):
    # Diccionario con los correos electrónicos de los clientes
    client_emails = {
    "CARNES Y ABARROTES A A A": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "GRANJERO FELIZ": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "LAURA MARIEL DIAZ ALVAREZ": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "COMERCIALIZADORA COLSEN": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "OSCAR EMILIO GARCIA GONZALEZ": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "GRUPO BRAVO NUÑO": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "DISTRIBUIDORA DE CARNES EL JAROCHO": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "PROCESADORA Y COMERCIALIZADORA CAMPEROS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "DISTRIBUIDORA DE CARNE LA ORIENTAL": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VENTA AL PUBLICO EN GENERAL - RB": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "MIGUEL ANGEL SANTIAGO HERNANDEZ": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "COMERCIALIZADORA MK DE SAN JUAN": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "BULMARO CASTILLO GUZMAN": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "CENTRO DE CARNES SAN ROBERTO": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "GRAFOLER": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "GRUPO DISTRIBUIDOR RANGEL GARDU&O": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "SIGMA FOODSERVICE COMERCIAL": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "AMERICA NANCY JAIMES OCAMPO": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VENTA AL PUBLICO EN GENERAL - SP": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "JESUS ALFONZO LOPEZ BERNAL": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "DISTRIBUIDORA DE PORCINOS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VIANSA ALIMENTOS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VENTA AL PUBLICO EN GENERAL - CF": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "ALFONSO ESPINDOLA SALDAÑA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "LA FRAILESCA CARNICERIA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "COMERCIALIZADORA DE CARNES LA HEROICA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "JOSE ANTONIO ARRONA CHIQUITO": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "ARCADIO LEDO BERISTAIN": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "SICARNEXS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "MARIO AARON PADILLA OLIVAS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "SOFIA LOMELI VARGAS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "MICAHELINA BUSTOS FIGUEROA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "ALAN DANIEL GARCIA RODRIGUEZ": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "JUAN ANGEL GONZALEZ TORRES": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "SDR ALIMENTOS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "JOSE ROBERTO COVARRUBIAS GARCIA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "AQUA TERRA IMPORTS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "OPERADORA FUTURAMA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    '"ALIMENTOS KARULY"': ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "DISTRIBUIDORA DUMY": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VENTA AL PUBLICO EN GENERAL - ML": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "ABASTECEDORA DE CARNICOS DEL SURESTE": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "ABASTO DE 4 CARNES": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VENTA AL PUBLICO EN GENERAL - MA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "ORGANIZACION REAL FOODS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "MANUEL JESUS CORONADO SOSA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "YESSICA CAICERO MURRIETA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "HURZEN FOODS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "PROCESADORA Y EMPACADORA DE CARNES SAN JOSE": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VICTOR HUGO ALDAMA DERAS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VENTAS AL PUBLICO EN GENERAL - DG": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "COMERCIALIZADORA DE CARNES JIVE DE JALISCO": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "PROCESADORA DE CARNES DON TIMO": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "CARNICOS DM": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "PROCESADORA DE ALIMENTOS OMEX": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "DELIMARKETS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "DISTRIBUIDORA DE CARNES SELECTAS MARCAF": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "ROGELIO HERNANDEZ SOTO": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "PROCESADORA DE CARNES EL JAROCHO": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "EMPACADORA ARIALY": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VENTA AL PUBLICO EN GENERAL - RS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "SALVADOR AURELIO EGURVIDE PIMENTEL": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "BEMEAT": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "ESCUDERO 1°": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "IDE FOODS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "PESCADOS Y MARISCOS DEL BAJIO": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "COMERCIALIZADORA MEXPORK": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "TRANSFORMACION CARNICA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "COMERCIALIZADORA PORCICOLA MEXICANA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VENTA AL PUBLICO EN GENERAL - RC": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "COMERCIALIZADORA DE CREMERIA Y CARNICOS DEL AHORRO": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VENTA AL PUBLICO EN GENERAL -AM": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VENTA AL PUBLICO EN GENERAL - IA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "ALFREDO CABAÑAS GUZMAN": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VENTA AL PUBLICO EN GENERAL - AB": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "BOCADOS JL": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "IMPULSORA DE BIENES ALAMEDA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "CARNES SELECTAS LA CAPITAL 6731": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VENTA AL PUBLICO EN GENERAL - IE": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "CESAR MOUREY GARCIA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "CARNES Y QUESOS REGIONALES": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "ADRIANA MARISELA MELLADO MENCIAS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "LA FRONTERA DISTRIBUIDORA DE CONGELADOS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "CARNES DELICIOSAS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "MN-BUSINESS SOLUTIONS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "CARLOS JESUS CANTO EUAN": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VENTA AL PUBLICO EN GENERAL - MC": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "PRODUCTOS CARNICOS AR": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "GIANT FOOD SERVICE": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "PRAC ALIMENTOS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "PROVEEDORA DE CARNES AGUASCALIENTES": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "FRIHMSA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "EMPACADORA Y PROCESADORA DE MONTERREY": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VENTA AL PUBLICO EN GENERAL - FL": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "SUPER CARNICERIAS HERNANDEZ": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "KAREY ALIMENTOS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "SERVI CARNES DE OCCIDENTE": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "MIGUEL ANGEL NAVARRO DIAZ": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "CARLOS DANIEL GARCIA GONZALEZ": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "LA CENTRAL SAN JUAN": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "GLENDA BELEM ROCHIN HERNANDEZ": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "F&J TRADING MEAT": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "EMPACADORA DE CARNES SELECTAS VICTORIA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "RUBEN GUERRERO CENDEJAS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "IMPORTADORA SERVICARNES": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "ALIMENTOS TAOR": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "ABAPROCAR": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "CARNICOS NACIONALES ZAPATA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "LACTEOS Y EMBUTIDOS DIEZ HERMANOS": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "SERVICIOS DE COMERCIALIZACION LOGISTICA DEL RIO": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "VIO ROCA COMERCIAL": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "TICI GENERAL PACKING": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "RAUL BOVIO GUERRERO": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "EMPACADORA D'CADENA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "PROTEINAS IMPORTADAS DE ALTA GAMA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "SERVICIOS ADUANALES HL": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "PRODUCTOS JAMEX": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "GRUPO QUERETARO CARNES": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "CHRISTIAN CAROLINA ESTRADA PONCE": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "COMERCIALIZADORA DE CARNES MOLOLO": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "PRODUCTOS NEZA": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"],
    "ROSARIO GONZALEZ SERRANO": ["mpleininger@isafoods.com", "mgomez@isafoods.com", "no-reply@isafoods.com"]
    }
    return client_emails.get(client_name, [])  # Retorna None si no encuentra el cliente


def clean_filename(client_name):
    cleaned_name = re.sub(r'[^a-zA-Z0-9-_]', '_', client_name)
    return cleaned_name


def send_email_with_attachment(to_emails, subject, html_body, attachment_paths, smtp_server, smtp_port, smtp_user,
                               smtp_password, client_name, image_paths=None):

    to_emails_str = ", ".join(to_emails)

    # Code for GMAIL
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = to_emails_str
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
        server.sendmail(smtp_user, to_emails, msg.as_string())
        server.quit()
        print(f"Email sent successfully for Client {client_name} to {to_emails}")
    except Exception as e:
        print(f"Failed to send email to {to_emails}: {e}")

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

    """


def csv_to_html_table(csv_path, name, moneda):
    # Leer el archivo CSV con pandas
    df = pd.read_csv(csv_path)
    df = df.fillna("")  # Llenar valores NaN con cadena vacía

    # Eliminar la columna 'Balance Vencido' si existe
    if 'Balance Vencido' in df.columns:
        df.drop(columns=['Balance Vencido'], inplace=True)

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

    # Establecer un tamaño fijo (ancho y alto) y controlar el texto en celdas
    cell_style = (
        'width: 160px; height: 40px; text-align: center; '
        'white-space: nowrap; overflow: hidden; text-overflow: ellipsis;'
    )
    html_table = html_table.replace('<th>', f'<th style="{cell_style}">')
    html_table = html_table.replace('<td>', f'<td style="{cell_style}">')

    # Título con la moneda
    title = f"<h3>Facturas de {name} en {'Pesos Mexicanos (MXN)' if moneda == 'MXN' else 'Dólares Americanos (USD)'}</h3>"

    # Devuelvo el HTML final con el título y la tabla estilizada
    html_final = title + html_table
    return html_final


def generate_file_name(client_name, currency):
    # Reemplazar espacios por "_"
    sanitized_client_name = (client_name.replace(" ", "_").replace("-", "_").replace('"', "_").replace("&", "_")
                             .replace("°", "_").replace("'", "_"))


    # Construir el nombre del archivo con moneda y sufijo
    file_name = f"{sanitized_client_name}_{currency}_invoices.csv"

    return file_name


def send_invoices_to_clients(organized_invoices, smtp_server, smtp_port, smtp_user, smtp_password, list_clients,
                             dict_clients_emails):
    client_fiscal = {
        "VENTA AL PUBLICO - CZ": "CARNICOS NACIONALES ZAPATA SA DE CV",
        "VENTA AL PUBLICO - ES": "",
        "VENTA AL PUBLICO - GH": "GLENDA BELEM ROCHIN HERNANDEZ",
        #"VENTA AL PUBLICO - IA": "ISA ALIMENTOS SA DE CV", # Cambio Name
        "VENTA AL PUBLICO EN GENERAL - IA": "ISA ALIMENTOS SA DE CV",
        "VENTA AL PUBLICO - LA": "LUIS ARREGUIN",
        "VENTA AL PUBLICO - OML": "OMAR LARA",
        "VENTA AL PUBLICO EN GENERAL  - AL": "HILDA ALICIA LARIOS HIDALGO",
        "VENTA AL PUBLICO EN GENERAL  - CF": "CARNICOS LA FORTUNA",
        "VENTA AL PUBLICO EN GENERAL  - FC": "MARTINEZ COBOS FRANCISCO",
        "VENTA AL PUBLICO EN GENERAL -  FCH": "",
        "VENTA AL PUBLICO EN GENERAL  - JE": "JORGE ESQUIVEL CASTRO",
        "VENTA AL PUBLICO EN GENERAL  - JR": "JASSO RODRIGUEZ CRUZ",
        "VENTA AL PUBLICO EN GENERAL  - MA": "JOSE MANUEL ALVARADO",
        "VENTA AL PUBLICO EN GENERAL  - MC": "MANUEL JESUS CORONADO SOSA",
        "VENTA AL PUBLICO EN GENERAL  - MN": "MIGUEL NAVARRO",
        "VENTA AL PUBLICO EN GENERAL  - MV": "MIGUEL VALVERDE",
        "VENTA AL PUBLICO EN GENERAL - AB": "ARTURO BARRADAS",
        "VENTA AL PUBLICO EN GENERAL - AG": "ABELARDO GONZALEZ",
        "VENTA AL PUBLICO EN GENERAL - AQ": "AQUA ELIT  SA DE CV",
        "VENTA AL PUBLICO EN GENERAL - CO": "",
        "VENTA AL PUBLICO EN GENERAL-  COP": "COPOCAR SA DE CV",
        "VENTA AL PUBLICO EN GENERAL - COV": "COVAGO CARNES SA DE CV",
        "VENTA AL PUBLICO EN GENERAL - EG": "ROSARIO GONZALEZ SERRANO",
        "VENTA AL PUBLICO EN GENERAL - FL": "FRANCISCO LOPEZ RAMIREZ",
        "VENTA AL PUBLICO EN GENERAL - GT": "GILDARDO TORRES",
        "VENTA AL PUBLICO EN GENERAL - IE": "IVAN ESTRADA ALVARADO",
        "VENTA AL PUBLICO EN GENERAL - JC": "JAVIER CASTILLO VIVEROS",
        "VENTA AL PUBLICO EN GENERAL - JM": "JONATHAN MARTINEZ BUSTOS",
        "VENTA AL PUBLICO EN GENERAL - JRL": "",
        "VENTA AL PUBLICO EN GENERAL - JS": "JUAN SANTIAGO SANTIAGO",
        "VENTA AL PUBLICO EN GENERAL - JV": "JOSE ALFREDO VELA BARQUERA",
        "VENTA AL PUBLICO EN GENERAL - JVG": "JOSUE VILLALPANDO GUZMAN",
        "VENTA AL PUBLICO EN GENERAL - MG": "MARBUSTELL GRUPO COMERCIAL SA DE CV",
        "VENTA AL PUBLICO EN GENERAL - ML": "MARIO LOPEZ VANZINI",
        "VENTA AL PUBLICO EN GENERAL - MS": "MIGUEL ANGEL SANTIAGO HERNANDEZ",
        "VENTA AL PUBLICO EN GENERAL - RB": "RAUL BOVIO GUERRERO",
        "VENTA AL PUBLICO EN GENERAL  - RC": "RAUL COSME",
        "VENTA AL PUBLICO EN GENERAL - RS": "EZEQUIEL VAZQUEZ SERRANO", #LA UNION
        "VENTA AL PUBLICO EN GENERAL - SP": "SAUL PEREZ",
        "VENTA AL PUBLICO EN GENERAL - VE": "VICENTE ESTRADA DOMINGUEZ",
        "VENTA AL PUBLICO EN GENERAL - YC": "CARNICERIA LA CABAÑA",
        "VENTA AL PUBLICO EN GENERAL -AM": "ADRIAN MONTIEL PEÑA",
        "VENTA AL PUBLICO GENERAL - OL": "OMAR LARA",
        #"VENTAS AL PUBLICO EN GENERAL - DG": "COMERCIALIZADORA MIZTLI & ELIZ S DE RL DE CV"
    }

    names_email = {
        "VENTA AL PUBLICO - CZ": "CARNICOS NACIONALES ZAPATA SA DE CV",
        "VENTA AL PUBLICO - ES": "ELIAS SERRANO GONZÁLEZ",
        "VENTA AL PUBLICO - GH": "GLENDA BELEM ROCHIN HERNANDEZ",
        #"VENTA AL PUBLICO - IA": "ISA ALIMENTOS SA DE CV", Cambio
        "VENTA AL PUBLICO EN GENERAL - IA": "ISA ALIMENTOS SA DE CV",
        "VENTA AL PUBLICO - LA": "LUIS ARREGUIN",
        "VENTA AL PUBLICO - OML": "OMAR LARA",
        "VENTA AL PUBLICO EN GENERAL  - AL": "HILDA ALICIA LARIOS HIDALGO",
        "VENTA AL PUBLICO EN GENERAL  - CF": "CARNICOS LA FORTUNA",
        "VENTA AL PUBLICO EN GENERAL  - FC": "MARTINEZ COBOS FRANCISCO",
        "VENTA AL PUBLICO EN GENERAL -  FCH": "FERNANDO CHAVEZ",
        "VENTA AL PUBLICO EN GENERAL  - JE": "JORGE ESQUIVEL CASTRO",
        "VENTA AL PUBLICO EN GENERAL  - JR": "JASSO RODRIGUEZ CRUZ",
        "VENTA AL PUBLICO EN GENERAL  - MA": "JOSE MANUEL ALVARADO",
        "VENTA AL PUBLICO EN GENERAL  - MC": "MANUEL JESUS CORONADO SOSA",
        "VENTA AL PUBLICO EN GENERAL  - MN": "MIGUEL NAVARRO",
        "VENTA AL PUBLICO EN GENERAL  - MV": "MIGUEL VALVERDE",
        "VENTA AL PUBLICO EN GENERAL - AB": "ARTURO BARRADAS",
        "VENTA AL PUBLICO EN GENERAL - AG": "ABELARDO GONZALEZ",
        "VENTA AL PUBLICO EN GENERAL - AQ": "AQUA ELIT SA DE CV",
        "VENTA AL PUBLICO EN GENERAL - CO": "COPROCAR",
        "VENTA AL PUBLICO EN GENERAL-  COP": "COPOCAR SA DE CV",
        "VENTA AL PUBLICO EN GENERAL - COV": "COVAGO CARNES SA DE CV",
        "VENTA AL PUBLICO EN GENERAL - EG": "ROSARIO GONZALEZ SERRANO",
        "VENTA AL PUBLICO EN GENERAL - FL": "FRANCISCO LOPEZ RAMIREZ",
        "VENTA AL PUBLICO EN GENERAL - GT": "GILDARDO TORRES",
        "VENTA AL PUBLICO EN GENERAL - IE": "IVAN ESTRADA ALVARADO",
        "VENTA AL PUBLICO EN GENERAL - JC": "JAVIER CASTILLO VIVEROS",
        "VENTA AL PUBLICO EN GENERAL - JM": "JONATHAN MARTINEZ BUSTOS",
        "VENTA AL PUBLICO EN GENERAL - JRL": "JAVIER RAMIREZ LOPEZ",
        "VENTA AL PUBLICO EN GENERAL - JS": "JUAN SANTIAGO SANTIAGO",
        "VENTA AL PUBLICO EN GENERAL - JV": "JOSE ALFREDO VELA BARQUERA",
        "VENTA AL PUBLICO EN GENERAL - JVG": "JOSUE VILLALPANDO GUZMAN",
        "VENTA AL PUBLICO EN GENERAL - MG": "MARBUSTELL GRUPO COMERCIAL SA DE CV",
        "VENTA AL PUBLICO EN GENERAL - ML": "MARIO LOPEZ VANZINI",
        "VENTA AL PUBLICO EN GENERAL - MS": "MIGUEL ANGEL SANTIAGO HERNANDEZ",
        "VENTA AL PUBLICO EN GENERAL - RB": "RAUL BOVIO GUERRERO",
        "VENTA AL PUBLICO EN GENERAL  - RC": "RAUL COSME",
        "VENTA AL PUBLICO EN GENERAL - RS": "LA UNION",
        "VENTA AL PUBLICO EN GENERAL - SP": "SAUL PEREZ",
        "VENTA AL PUBLICO EN GENERAL - VE": "VICENTE ESTRADA DOMINGUEZ",
        "VENTA AL PUBLICO EN GENERAL - YC": "CARNICERIA LA CABAÑA",
        "VENTA AL PUBLICO EN GENERAL -AM": "ADRIAN MONTIEL PEÑA",
        "VENTA AL PUBLICO GENERAL - OL": "OMAR LARA",
        "VENTAS AL PUBLICO EN GENERAL - DG": "DON GATO"
    }

    for client, invoices_list in organized_invoices.items():
        client_emails = dict_clients_emails.get(client, [])
        #client_emails = get_client_emails(client)

        name_fiscal = None
        names = []
        names.append(client)
        signal = False
        for key, value in client_fiscal.items():
            if client == key:
                if value in list_clients:
                    names.append(value)
                    name_fiscal = names_email.get(key)

                else:
                    name_fiscal = names_email.get(key)
            if client == value:
                if key in list_clients:
                    signal = True
                else:
                    pass
        if signal:
            continue

        if client_emails:

            # Leer cada archivo CSV para calcular los balances vencidos
            attachment_paths = []

            body_html = f"""
            <html>
                <body>
                    <img src="cid:image_1" alt="Logo" style="width:150px; margin-top: 20px;">
                    <h3 style="text-align: left;">ESTADO DE CUENTA</h3>
                    <p>Estimado <strong>{name_fiscal if name_fiscal else client}</strong>,</p>
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
                    MXN = False
                    USD = False

                    file_name_client = generate_file_name(name, currency)
                    file_path = os.path.join("output", file_name_client)

                    if not os.path.exists(file_path):
                        print(f"Archivo {file_path} no encontrado. Continuando con la siguiente iteración.")
                        continue

                    # Leer el CSV para obtener el balance vencido
                    try:
                        df = pd.read_csv(file_path)
                        if "Balance Vencido" in df.columns:
                            valores_no_nulos = df["Balance Vencido"].dropna()

                            if not valores_no_nulos.empty:  # Verificar si hay valores después de eliminar NaN
                                last_balance = float(valores_no_nulos.iloc[-1])  # Obtener el último valor válido
                            else:
                                last_balance = 0  # Si no hay valores, asignar 0

                        else:
                            last_balance = 0

                        if currency == "MXN":
                            balance_vencido_mxn += last_balance
                            MXN = True
                        elif currency == "USD":
                            balance_vencido_usd += last_balance
                            USD = True
                    except Exception as e:
                        print(f"Error reading {file_path}: {e}")
                        continue

                    # Generar tabla HTML para el archivo y agregar al cuerpo del correo
                    attachment_paths.append(file_path)

                    if MXN:
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
                    if USD:
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
            subject = f"Estado de Cuenta para {name_fiscal if name_fiscal else client} - {format_date}"


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
                <p><br /><br /><strong>Saludos</strong>.</p>
            """
            body_html += f"""
                <img src="cid:image_2" alt="Second Image" style="width:676px; margin-top: 20px; margin-bottom: 0;">
                </body>
                </html>
                """

            # Enviar correo con tablas y archivos adjuntos
            send_email_with_attachment(client_emails, subject, body_html, attachment_paths, smtp_server, smtp_port,
                                       smtp_user, smtp_password, client, image_paths)
        else:
            print(f"Email not found for client {client}")
