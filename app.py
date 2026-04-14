#!/usr/bin/env python3
"""Sales Order V3 — BPRO/PO shortfall tracker for XpertPack.

Reads the 'Final Tracker all Customer SO' tab in ONE API call
instead of cycling through 87 customers one at a time.
"""

from flask import Flask, jsonify, send_from_directory
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import os

app = Flask(__name__, static_folder='static')

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
CREDS_FILE = os.path.join(os.path.dirname(__file__), 'credentials.json')
SPREADSHEET_ID = '1mo8yEcY7V6lMMpDpjHv6MTxfwLXvsQCkgvJk1MXfrhM'

# Column indices in 'Final Tracker all Customer SO' (row 2 = headers)
COL_CUSTOMER = 0       # Customer Name
COL_CUST_ITEM = 1      # Customer Item Code
COL_ITEM_CODE = 2      # Item Code
COL_TYPE = 4           # Item type (header says "Foam" but holds type)
COL_SO = 5             # Sales Order No.
COL_PENDING = 11       # SUM of Final Pending Qty
COL_FG = 12            # FG (Finished Goods stock)
COL_BPRO_NO = 13       # BPRO. NO.
COL_PO_NO = 22         # Purchase Order No

# Types that need BPRO (manufactured in-house corrugated/lid items)
BPRO_TYPES = {'corr', 'Lid', 'Die', 'Sleeve'}
# Types that need PO (bought-out items)
PO_TYPES = {'Plastic', 'Foam', 'Consumable', 'VCI', 'ESD', 'LDPE/HDPE',
            'wrap', 'tapes', 'straps', 'stickers', 'Chemical', 'Metal'}


def get_service():
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)


def safe_float(v):
    try:
        return float(str(v).replace(',', '')) if v and str(v).strip() else 0
    except:
        return 0


def col(row, idx):
    """Safely get column value from row."""
    return row[idx].strip() if len(row) > idx and row[idx] else ''


@app.route('/')
def index():
    return send_from_directory('static', 'index.html')


@app.route('/api/shortfalls')
def api_shortfalls():
    """Scan ALL customers in ONE read and return items needing BPRO or PO."""
    service = get_service()

    # Single API call — reads all ~4400 rows at once
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range="'Final Tracker all Customer SO'!A3:AG10000"
    ).execute()
    rows = result.get('values', [])

    needs_bpro = []
    needs_po = []
    customers_seen = set()

    for row in rows:
        item_code = col(row, COL_ITEM_CODE)
        if not item_code:
            continue

        customer = col(row, COL_CUSTOMER)
        if not customer:
            continue

        customers_seen.add(customer)
        item_type = col(row, COL_TYPE)
        pending = safe_float(col(row, COL_PENDING))
        fg = safe_float(col(row, COL_FG))
        bpro_no = col(row, COL_BPRO_NO)
        po_no = col(row, COL_PO_NO)
        net_need = pending - fg

        if pending <= 0:
            continue

        entry = {
            'customer': customer,
            'customer_item': col(row, COL_CUST_ITEM),
            'item_code': item_code,
            'type': item_type,
            'so': col(row, COL_SO),
            'pending': pending,
            'fg': fg,
            'net_need': net_need,
        }

        if item_type in BPRO_TYPES and net_need > 0 and not bpro_no:
            needs_bpro.append(entry)

        if item_type in PO_TYPES and net_need > 0 and not po_no:
            needs_po.append(entry)

    return jsonify({
        'bpro': needs_bpro,
        'po': needs_po,
        'summary': {
            'bpro_count': len(needs_bpro),
            'po_count': len(needs_po),
            'customers_scanned': len(customers_seen),
        }
    })


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5007, debug=True)
