from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP


def organize_invoices_by_client(invoices):
    organized_data = defaultdict(list)
    for invoice in invoices:
        client_name = invoice.get("ClientName", "Unknown")
        organized_data[client_name].append(invoice)
    return organized_data


def add_payment_details_to_invoices(organized_invoices, token, fetch_payment_details_func):
    for invoices in organized_invoices.values():
        for invoice in invoices:
            invoice_id = invoice.get("ID")
            if invoice_id:
                invoice["PaymentDetails"] = fetch_payment_details_func(invoice_id, token)


def format_decimal(value, decimal_places=2):
    return Decimal(value).quantize(Decimal(f"1.{'0' * decimal_places}"), rounding=ROUND_HALF_UP)

