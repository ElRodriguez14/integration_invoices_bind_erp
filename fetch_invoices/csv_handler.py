import os
import csv
import locale

from .data_processing import format_decimal
from decimal import Decimal
from datetime import datetime


def clean_filename(client_name):
    return "".join(c if c.isalnum() else "_" for c in client_name)


def format_date_to_yyyy_mm_dd(date_string):
    """Convierte una fecha a formato YYYY-MM-DD."""
    formats = [
        "%Y-%m-%d",                # Solo fecha
        "%Y-%m-%dT%H:%M:%S.%f",    # Fecha con fracciones de segundo
        "%Y-%m-%dT%H:%M:%S"        # Fecha con tiempo completo, sin fracciones
    ]
    for fmt in formats:
        try:
            date_obj = datetime.strptime(date_string, fmt)
            return date_obj.strftime("%Y-%m-%d")
        except ValueError:
            continue
    print(f"Error: fecha '{date_string}' no válida. Se omitirá.")
    return None


def format_date_to_text(date_string):
    """
    Convierte una fecha en formato DD/MM/YYYY al formato 'Month Day, Year'.
    """
    formats = [
        "%Y-%m-%d",                # Solo fecha
        "%Y-%m-%dT%H:%M:%S.%f",    # Fecha con fracciones de segundo
        "%Y-%m-%dT%H:%M:%S"        # Fecha con tiempo completo, sin fracciones
    ]
    for fmt in formats:
        try:
            date_obj = datetime.strptime(date_string, fmt)
            locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
            formatted_date = date_obj.strftime("%B %d %Y")
            formatted_date = formatted_date.capitalize()
            locale.setlocale(locale.LC_TIME, '')
            return formatted_date
        except ValueError:
            continue
    print(f"Error: fecha '{date_string}' no válida. Se omitirá.")
    return None


def calculate_days_overdue(expiration_date):
    """Calcula los días vencidos respecto a la fecha de hoy."""
    try:
        if isinstance(expiration_date, str):
            expiration_date = datetime.strptime(expiration_date, "%Y-%m-%d")  # Asegúrate de usar el formato correcto
        today = datetime.now()
        days_overdue = (today - expiration_date).days
        return days_overdue
    except (ValueError, TypeError):
        return 0


def export_invoices_to_csv(organized_invoices, output_dir="output"):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)  # Crear la carpeta si no existe

    for client, currencies in organized_invoices.items():
        for currency, invoices_list in currencies.items():
            file_name = f"{clean_filename(client)}_{currency}_invoices.csv"
            file_path = os.path.join(output_dir, file_name)

            # Asegurarse de que hay datos para escribir
            if not invoices_list:
                print(f"No invoices to write for client {client}.")
                continue

            sorted_invoices = sorted(
                invoices_list,
                key=lambda x: format_date_to_yyyy_mm_dd(x.get("Date"))
            )

            with open(file_path, mode="w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                # Escribir el encabezado
                writer.writerow(["Fecha", "Descripcion", "Factura", "PO", "Fecha Vencimiento", "Dias Vencidos",
                                 "Total", "Balance"])

                accumulated_balance = Decimal("0")
                # Escribir los datos
                for invoice in sorted_invoices:
                    # Factura
                    invoice_date = format_date_to_text(invoice.get("Date"))
                    invoice_id = invoice.get("Number")
                    purchase_order = invoice.get("PurchaseOrder")
                    expiration_date = format_date_to_yyyy_mm_dd(invoice.get("ExpirationDate"))
                    credit_notes = format_decimal(Decimal(invoice.get("CreditNotes", 0)))
                    total = format_decimal(Decimal(invoice.get("Total", 0)))

                    days_overdue = calculate_days_overdue(expiration_date)
                    expiration_date = format_date_to_text(invoice.get("ExpirationDate"))


                    if invoice["PaymentDetails"] == []:
                        accumulated_balance += total
                        formatted_accumulated_balance = format_decimal(accumulated_balance)
                        writer.writerow([
                            invoice_date, "Factura", int(invoice_id or 0), int(purchase_order or 0),
                            expiration_date, int(days_overdue), total, formatted_accumulated_balance
                        ])
                    else:
                        # Agregar fila de factura sin cambios
                        accumulated_balance += total
                        formatted_accumulated_balance = format_decimal(accumulated_balance)
                        writer.writerow([
                            invoice_date, "Factura", int(invoice_id or 0), int(purchase_order or 0),
                            expiration_date, int(days_overdue), total, formatted_accumulated_balance
                        ])

                    if credit_notes > 0:
                        accumulated_balance -= credit_notes
                        formatted_balance_nt = format_decimal(accumulated_balance)

                        # Agregar la fila de credit notes
                        writer.writerow([invoice_date, "Nota Credito", "", "", "", "",
                                         f"({credit_notes})", formatted_balance_nt])

                    # Agregar los pagos
                    if "PaymentDetails" in invoice:

                        sorted_payments = sorted(
                            invoice["PaymentDetails"],
                            key=lambda x: format_date_to_yyyy_mm_dd(x.get("ApplicationDate"))
                        )

                        for payment in sorted_payments:
                            if currency == "USD":
                                amount = Decimal(payment.get("Amount")) / Decimal(payment.get("ExchangeRate"))
                            else:
                                amount = Decimal(payment.get("Amount", 0))

                            application_date = format_date_to_text(payment.get("ApplicationDate"))

                            accumulated_balance -= amount
                            formatted_amount_payment = format_decimal(amount)
                            formatted_remaining = format_decimal(accumulated_balance)

                            # Agregar la fila de pago
                            writer.writerow([application_date, "Pago", "", "", "", "",
                                             f"({formatted_amount_payment})", formatted_remaining])

            print(f"Invoices for client {client} in {currency} exported to {file_path}")

    print("Finish Export")
