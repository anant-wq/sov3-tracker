#!/usr/bin/env python3
"""Google Sheets reader/writer for Sales Order V3 sheet."""

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import os

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CREDS_FILE = os.path.join(os.path.dirname(__file__), 'credentials.json')
SPREADSHEET_ID = '1mo8yEcY7V6lMMpDpjHv6MTxfwLXvsQCkgvJk1MXfrhM'


def get_service():
    """Authenticate and return a Sheets API service."""
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)


def list_sheets():
    """List all sheet/tab names in the spreadsheet."""
    service = get_service()
    result = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    return [s['properties']['title'] for s in result['sheets']]


def read_sheet(sheet_name='ERP Dump', range_str=None, rows=None):
    """Read data from a sheet tab. Optionally limit rows."""
    service = get_service()
    if range_str:
        full_range = f"'{sheet_name}'!{range_str}"
    else:
        full_range = sheet_name
    data = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=full_range
    ).execute()
    values = data.get('values', [])
    if rows:
        return values[:rows]
    return values


def read_all(sheet_name='ERP Dump'):
    """Read all data from a sheet tab as list of dicts (header = first row)."""
    values = read_sheet(sheet_name)
    if len(values) < 2:
        return []
    headers = values[0]
    return [dict(zip(headers, row)) for row in values[1:]]


def write_cell(sheet_name, cell, value):
    """Write a value to a specific cell (e.g. 'A1')."""
    service = get_service()
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!{cell}",
        valueInputOption='USER_ENTERED',
        body={'values': [[value]]}
    ).execute()
    return f"Written '{value}' to {cell} in {sheet_name}"


def write_range(sheet_name, start_cell, values):
    """Write a 2D list of values starting from start_cell."""
    service = get_service()
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!{start_cell}",
        valueInputOption='USER_ENTERED',
        body={'values': values}
    ).execute()
    return f"Written {len(values)} rows to {sheet_name} starting at {start_cell}"


if __name__ == '__main__':
    print("=== Connecting to Google Sheets ===")

    sheets = list_sheets()
    print(f"\nSheet tabs ({len(sheets)}): {sheets}")

    print(f"\n=== First 3 rows from '{sheets[0]}' ===")
    rows = read_sheet(sheets[0], rows=3)
    for i, row in enumerate(rows):
        print(f"Row {i}: {row[:5]}...")
