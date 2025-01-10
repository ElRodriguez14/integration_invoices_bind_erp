import requests
import time

def fetch_invoices(api_url, token):
    headers = {"Authorization": f"Bearer {token}"}
    invoices = []
    next_link = api_url

    while next_link:
        print(f"Fetching data from: {next_link}")
        response = requests.get(next_link, headers=headers)
        if response.status_code == 200:
            data = response.json()
            invoices.extend(data.get("value", []))
            next_link = data.get("nextLink")
        else:
            print(f"Error {response.status_code}: {response.text}")
            break
    return invoices


def fetch_payment_details(invoice_id, token):
    payment_url = f"http://api.bind.com.mx/api/Invoices/Payment/{invoice_id}"
    headers = {"Authorization": f"Bearer {token}"}

    while True:
        response = requests.get(payment_url, headers=headers)
        if response.status_code == 200:
            return response.json().get("value", [])
        elif response.status_code == 429:
            print(f"Rate limit exceeded for Invoice ID {invoice_id}. Retrying...")
            time.sleep(5)
        else:
            print(f"Error {response.status_code}: {response.text}")
            return []
