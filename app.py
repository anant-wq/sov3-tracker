#!/usr/bin/env python3
"""Sales Order V3 — BPRO/PO shortfall tracker for XpertPack."""

from flask import Flask, jsonify, send_from_directory
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import time, json, os

app = Flask(__name__, static_folder='static')

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CREDS_FILE = os.path.join(os.path.dirname(__file__), 'credentials.json')
SPREADSHEET_ID = '1mo8yEcY7V6lMMpDpjHv6MTxfwLXvsQCkgvJk1MXfrhM'


def get_service():
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)


def safe_float(v):
    try:
        return float(v.replace(',', '')) if v and v.strip() else 0
    except:
        return 0


@app.route('/')
def index():
    return send_from_directory('static', 'index.html')


@app.route('/api/customers')
def api_customers():
    """Get list of all customers."""
    service = get_service()
    data = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range="'ERP Dump'!C2:C10000"
    ).execute()
    customers = sorted(set(r[0] for r in data.get('values', []) if r and r[0].strip()))
    return jsonify(customers)


@app.route('/api/customer/<path:customer_name>')
def api_customer_data(customer_name):
    """Get tracker data for a single customer."""
    service = get_service()
    # Write customer to B1
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range="'Final Tracker Customer Wise SO'!B1",
        valueInputOption='USER_ENTERED',
        body={'values': [[customer_name]]}
    ).execute()
    time.sleep(1.5)

    # Read filtered data
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range="'Final Tracker Customer Wise SO'!A3:AG500"
    ).execute()
    rows = result.get('values', [])

    items = []
    for row in rows:
        if len(row) < 12:
            continue
        if not (row[2].strip() if len(row) > 2 else ''):
            continue

        pending = safe_float(row[11]) if len(row) > 11 else 0
        fg = safe_float(row[12]) if len(row) > 12 else 0
        net_need = pending - fg

        items.append({
            'customer': row[0] if len(row) > 0 else '',
            'customer_item': row[1] if len(row) > 1 else '',
            'item_code': row[2] if len(row) > 2 else '',
            'description': row[3] if len(row) > 3 else '',
            'type': row[4] if len(row) > 4 else '',
            'so': row[5] if len(row) > 5 else '',
            'so_date': row[6] if len(row) > 6 else '',
            'customer_po': row[7] if len(row) > 7 else '',
            'qty': safe_float(row[9]) if len(row) > 9 else 0,
            'invoiced': safe_float(row[10]) if len(row) > 10 else 0,
            'pending': pending,
            'fg': fg,
            'bpro_no': row[13].strip() if len(row) > 13 else '',
            'bpro_qty': safe_float(row[14]) if len(row) > 14 else 0,
            'bpro_date': row[15].strip() if len(row) > 15 else '',
            'ipro_qty': safe_float(row[16]) if len(row) > 16 else 0,
            'ipro_no': row[17].strip() if len(row) > 17 else '',
            'po_no': row[22].strip() if len(row) > 22 else '',
            'po_qty': safe_float(row[23]) if len(row) > 23 else 0,
            'action_point': safe_float(row[26]) if len(row) > 26 else 0,
            'net_need': net_need,
        })

    return jsonify(items)


@app.route('/api/shortfalls')
def api_shortfalls():
    """Scan ALL customers and return items needing BPRO or PO."""
    service = get_service()

    # Get customer list
    data = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range="'ERP Dump'!C2:C10000"
    ).execute()
    customers = sorted(set(r[0] for r in data.get('values', []) if r and r[0].strip()))

    needs_bpro = []
    needs_po = []

    for cust in customers:
        try:
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range="'Final Tracker Customer Wise SO'!B1",
                valueInputOption='USER_ENTERED',
                body={'values': [[cust]]}
            ).execute()
            time.sleep(1.5)

            result = service.spreadsheets().values().get(
                spreadsheetId=SPREADSHEET_ID,
                range="'Final Tracker Customer Wise SO'!A3:AG500"
            ).execute()
            rows = result.get('values', [])

            for row in rows:
                if len(row) < 12:
                    continue
                if not (row[2].strip() if len(row) > 2 else ''):
                    continue

                item_type = row[4] if len(row) > 4 else ''
                pending = safe_float(row[11]) if len(row) > 11 else 0
                fg = safe_float(row[12]) if len(row) > 12 else 0
                bpro_no = row[13].strip() if len(row) > 13 else ''
                po_no = row[22].strip() if len(row) > 22 else ''
                net_need = pending - fg

                if pending <= 0:
                    continue

                entry = {
                    'customer': cust,
                    'customer_item': row[1] if len(row) > 1 else '',
                    'item_code': row[2] if len(row) > 2 else '',
                    'type': item_type,
                    'so': row[5] if len(row) > 5 else '',
                    'pending': pending,
                    'fg': fg,
                    'net_need': net_need,
                }

                if item_type in ('corr', 'Lid') and net_need > 0 and not bpro_no:
                    needs_bpro.append(entry)

                if item_type in ('Plastic', 'Foam', 'Consumable') and net_need > 0 and not po_no:
                    needs_po.append(entry)

        except Exception as e:
            print(f"Error for {cust}: {e}")

    return jsonify({
        'bpro': needs_bpro,
        'po': needs_po,
        'summary': {
            'bpro_count': len(needs_bpro),
            'po_count': len(needs_po),
            'customers_scanned': len(customers),
        }
    })


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5003, debug=True)
