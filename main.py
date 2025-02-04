from fetch_invoices.api import fetch_invoices, fetch_payment_details
from fetch_invoices.data_processing import organize_invoices_by_client_and_currency, add_payment_details_to_invoices
from fetch_invoices.csv_handler import export_invoices_to_csv
from fetch_invoices.email_sender import send_invoices_to_clients
from config.settings import SMTP_CONFIG, API_CONFIG

if __name__ == "__main__":
    print("Fetching invoices...")
    invoices = fetch_invoices(API_CONFIG["url"], API_CONFIG["token"])
    print(f"Total invoices fetched: {len(invoices)}")

    organized_invoices = organize_invoices_by_client_and_currency(invoices)
    print(f"Invoices organized by ClientName: {len(organized_invoices)} clients")

    print("Adding payment details...")
    #add_payment_details_to_invoices(organized_invoices, API_CONFIG["token"], fetch_payment_details)

    print("Exporting invoices to CSV...")
    output_dir = "output"
    #export_invoices_to_csv(organized_invoices, output_dir)

    print("Sending invoices to clients...")

    smtp_server = "smtp.gmail.com"  # Cambia a tu servidor SMTP
    smtp_port = 587  # Usualmente 587 para TLS
    smtp_user = "integrationbind@gmail.com"  # Tu correo electrónico
    smtp_password = "rzsz lzyn rxxx yqsy"  # Tu contraseña

    send_invoices_to_clients(organized_invoices, smtp_server, smtp_port, smtp_user, smtp_password)


