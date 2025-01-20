from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP


def organize_invoices_by_client_and_currency(invoices):
    excluded_ids = {1031, 538, 509, 336, 262}
    organized_data = defaultdict(lambda: defaultdict(list))
    for invoice in invoices:
        invoice_id = invoice["Number"]
        if invoice_id in excluded_ids:
            continue
        client_name = invoice.get("ClientName", "Unknown")
        exchange_rate = invoice.get("ExchangeRate", 0)
        currency = "MXN" if exchange_rate == 1 else "USD"
        organized_data[client_name][currency].append(invoice)

    for client in organized_data:
        print(client)

    return organized_data



def add_payment_details_to_invoices(organized_invoices, token, fetch_payment_details_func):
    for client, currencies in organized_invoices.items():
        for currency, invoices in currencies.items():
            for invoice in invoices:
                invoice_id = invoice.get("ID")
                if invoice_id:
                    invoice["PaymentDetails"] = fetch_payment_details_func(invoice_id, token)


def format_decimal(value, decimal_places=2):
    return Decimal(value).quantize(Decimal(f"1.{'0' * decimal_places}"), rounding=ROUND_HALF_UP)

