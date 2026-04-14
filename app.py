#!/usr/bin/env python3
"""Sales Order V3 — BPRO/PO/IPRO shortfall tracker for XpertPack.

Two data sources, each read in a single API call:
1. SO V3 sheet — 'Final Tracker all Customer SO' tab (Sales Orders)
2. Monthly Plan sheet — 'Auto Working Sheet' tab (Monthly Plans with BOM)
"""

from flask import Flask, jsonify, send_from_directory
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import os

app = Flask(__name__, static_folder='static')

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
CREDS_FILE = os.path.join(os.path.dirname(__file__), 'credentials.json')

# Sheet IDs
SO_SHEET_ID = '1mo8yEcY7V6lMMpDpjHv6MTxfwLXvsQCkgvJk1MXfrhM'
MP_SHEET_ID = '1orR4_YhWN-jEvAHZRn2e_crIhRWbemE9BSEvzPJlQYI'

# --- SO V3 column indices ('Final Tracker all Customer SO') ---
SO_CUSTOMER = 0
SO_CUST_ITEM = 1
SO_ITEM_CODE = 2
SO_TYPE = 4
SO_SO = 5
SO_PENDING = 11
SO_FG = 12
SO_BPRO_NO = 13
SO_PO_NO = 22

# --- Monthly Plan column indices ('Auto Working Sheet') ---
MP_NAME = 0          # MP number
MP_CUSTOMER = 1      # Customer
MP_ITEM_CODE = 3     # Item Code (parent)
MP_BOM_LEVEL = 5     # BOM Level (empty=top/standalone, 1, 2)
MP_TYPE = 6          # Item Type
MP_CUST_ITEM = 7     # Customer Item Code
MP_CHILD_ITEM = 24   # Item Code (actual child code)
MP_FINAL_PENDING = 35
MP_FG = 36
MP_BPRO_NO = 37
MP_BPRO_QTY = 38
MP_IPRO_QTY = 39
MP_IPRO_NO = 40
MP_IPRO_BOARD_QTY = 41
MP_IPRO_PMS_QTY = 42
MP_WIP = 43
MP_THERMO_IPRO = 44
MP_THERMO_QTY = 45
MP_PO_NO = 46
MP_PO_QTY = 47
MP_PP_IPRO_NO = 48
MP_PP_IPRO_QTY = 49
MP_STITCHING = 50
MP_ACTION = 51

BPRO_TYPES = {'corr', 'Lid', 'Die', 'Sleeve'}
PO_TYPES = {'Plastic', 'Foam', 'Consumable', 'VCI', 'ESD', 'LDPE/HDPE',
            'wrap', 'tapes', 'straps', 'stickers', 'Chemical', 'Metal',
            'pp', 'PET'}


def get_service():
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)


def sf(v):
    """Safe float conversion."""
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
    """SO V3: Scan ALL customers in ONE read."""
    service = get_service()

    result = service.spreadsheets().values().get(
        spreadsheetId=SO_SHEET_ID,
        range="'Final Tracker all Customer SO'!A3:AG10000"
    ).execute()
    rows = result.get('values', [])

    needs_bpro = []
    needs_po = []
    customers_seen = set()

    for row in rows:
        item_code = col(row, SO_ITEM_CODE)
        if not item_code:
            continue
        customer = col(row, SO_CUSTOMER)
        if not customer:
            continue

        customers_seen.add(customer)
        item_type = col(row, SO_TYPE)
        pending = sf(col(row, SO_PENDING))
        fg = sf(col(row, SO_FG))
        bpro_no = col(row, SO_BPRO_NO)
        po_no = col(row, SO_PO_NO)
        net_need = pending - fg

        if pending <= 0:
            continue

        entry = {
            'customer': customer,
            'customer_item': col(row, SO_CUST_ITEM),
            'item_code': item_code,
            'type': item_type,
            'so': col(row, SO_SO),
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


@app.route('/api/monthly-plan')
def api_monthly_plan():
    """Monthly Plan: Read Auto Working Sheet in one call.

    Returns parent items (composite) with their children,
    BPRO/IPRO/PO data, and shortfall flag (Action < 0).
    """
    service = get_service()

    result = service.spreadsheets().values().get(
        spreadsheetId=MP_SHEET_ID,
        range="'Auto Working Sheet'!A2:AZ20000"
    ).execute()
    rows = result.get('values', [])

    items = []
    customers_seen = set()
    shortfall_count = 0
    needs_bpro = 0
    needs_ipro = 0
    needs_po = 0

    for row in rows:
        item_type = col(row, MP_TYPE)
        customer = col(row, MP_CUSTOMER)
        if not customer:
            continue

        # Use child item code if available, else parent
        item_code = col(row, MP_CHILD_ITEM) or col(row, MP_ITEM_CODE)
        if not item_code:
            continue

        customers_seen.add(customer)
        bom_level = col(row, MP_BOM_LEVEL)
        final_pending = sf(col(row, MP_FINAL_PENDING))
        fg = sf(col(row, MP_FG))
        bpro_no = col(row, MP_BPRO_NO)
        bpro_qty = sf(col(row, MP_BPRO_QTY))
        ipro_no = col(row, MP_IPRO_NO)
        ipro_qty = sf(col(row, MP_IPRO_QTY))
        po_no = col(row, MP_PO_NO)
        po_qty = sf(col(row, MP_PO_QTY))
        action = sf(col(row, MP_ACTION))
        net_need = final_pending - fg
        is_shortfall = action < 0 and final_pending > 0
        is_composite = item_type == 'composite'

        if final_pending <= 0 and not is_composite:
            continue

        entry = {
            'mp': col(row, MP_NAME),
            'customer': customer,
            'customer_item': col(row, MP_CUST_ITEM),
            'item_code': item_code,
            'parent_item': col(row, MP_ITEM_CODE),
            'type': item_type,
            'bom_level': bom_level,
            'is_composite': is_composite,
            'final_pending': final_pending,
            'fg': fg,
            'net_need': net_need,
            'bpro_no': bpro_no,
            'bpro_qty': bpro_qty,
            'ipro_no': ipro_no,
            'ipro_qty': ipro_qty,
            'po_no': po_no,
            'po_qty': po_qty,
            'wip': sf(col(row, MP_WIP)),
            'action': action,
            'is_shortfall': is_shortfall,
        }

        items.append(entry)

        if not is_composite and final_pending > 0:
            if is_shortfall:
                shortfall_count += 1
            if item_type in BPRO_TYPES and net_need > 0 and not bpro_no:
                needs_bpro += 1
            if item_type in BPRO_TYPES and net_need > 0 and not ipro_no:
                needs_ipro += 1
            if item_type in PO_TYPES and net_need > 0 and not po_no:
                needs_po += 1

    return jsonify({
        'items': items,
        'summary': {
            'total_items': len(items),
            'shortfall_count': shortfall_count,
            'needs_bpro': needs_bpro,
            'needs_ipro': needs_ipro,
            'needs_po': needs_po,
            'customers': len(customers_seen),
        }
    })


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5007, debug=True)
