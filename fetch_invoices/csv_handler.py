import os
import csv
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


def calculate_days_overdue(expiration_date):
    """Calcula los días vencidos respecto a la fecha de hoy."""
    try:
        today = datetime.now()
        days_overdue = (today - expiration_date).days
        return max(days_overdue, 0)
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
                writer.writerow(["Fecha", "Descripcion", "Invoice_id", "PO", "Fecha_Vencimiento", "Empleado",
                                 "Cuenta", "Total", "Balance", "Dias Vencidos"])

                accumulated_balance = Decimal("0")
                # Escribir los datos
                for invoice in sorted_invoices:
                    # Factura
                    invoice_date = format_date_to_yyyy_mm_dd(invoice.get("Date"))
                    invoice_id = invoice.get("Number")
                    purchase_order = invoice.get("PurchaseOrder")
                    expiration_date = format_date_to_yyyy_mm_dd(invoice.get("ExpirationDate"))
                    credit_notes = format_decimal(Decimal(invoice.get("CreditNotes", 0)))
                    total = format_decimal(Decimal(invoice.get("Total", 0)))
                    days_overdue = calculate_days_overdue(expiration_date)

                    values_pending = format_decimal(Decimal(invoice.get("Payments", 0))) + credit_notes
                    remaining_amount = format_decimal(total - values_pending)

                    accumulated_balance += remaining_amount

                    if invoice["PaymentDetails"] == []:
                        writer.writerow([
                            invoice_date, "Factura (Pendiente)", invoice_id, purchase_order,
                            expiration_date, "", "", total, accumulated_balance, days_overdue
                        ])
                    else:
                        # Agregar fila de factura sin cambios
                        writer.writerow([
                            invoice_date, "Factura", invoice_id, purchase_order,
                            expiration_date, "", "", total, accumulated_balance, days_overdue
                        ])

                    if credit_notes > 0:
                        # Agregar la fila de credit notes
                        writer.writerow([invoice_date, "Nota Credito", "", "", "",
                                         "", "", credit_notes])

                    # Agregar los pagos
                    if "PaymentDetails" in invoice:
                        for payment in invoice["PaymentDetails"]:
                            application_date = format_date_to_yyyy_mm_dd(payment.get("ApplicationDate"))
                            employee = payment.get("Employee")
                            account = payment.get("Account")
                            amount = format_decimal(Decimal(payment.get("Amount", 0)))

                            # Agregar la fila de pago
                            writer.writerow([application_date, "Pago", "", "", "",
                                             employee, account, amount])

            print(f"Invoices for client {client} in {currency} exported to {file_path}")

    print("Finish Export")
