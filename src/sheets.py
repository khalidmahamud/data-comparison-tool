"""
Google Sheets integration for IHADIS Data Comparison Tool.

Uses service account authentication to:
- Import data from Google Sheets
- Export data back to Google Sheets
"""
import os
import re
from pathlib import Path
from typing import Optional, List, Dict, Tuple

import pandas as pd

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False


# Google Sheets API scopes
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.readonly'
]


def get_credentials():
    """Get Google service account credentials."""
    if not GSPREAD_AVAILABLE:
        raise ImportError("gspread is not installed. Run: pip install gspread")

    creds_path = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS', 'service_account.json')

    if not Path(creds_path).exists():
        raise FileNotFoundError(f"Service account file not found: {creds_path}")

    credentials = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    return credentials


def get_client():
    """Get authenticated gspread client."""
    credentials = get_credentials()
    return gspread.authorize(credentials)


def extract_sheet_id(url_or_id: str) -> str:
    """
    Extract Google Sheet ID from URL or return as-is if already an ID.

    Supports formats:
    - https://docs.google.com/spreadsheets/d/SHEET_ID/edit
    - https://docs.google.com/spreadsheets/d/SHEET_ID/edit#gid=0
    - SHEET_ID (direct ID)
    """
    # Pattern to match Google Sheets URL
    pattern = r'/spreadsheets/d/([a-zA-Z0-9-_]+)'
    match = re.search(pattern, url_or_id)

    if match:
        return match.group(1)

    # Assume it's already an ID if no match
    return url_or_id


def get_sheet_info(url_or_id: str) -> Dict:
    """
    Get information about a Google Sheet.

    Returns:
        dict with title, sheet names, and row counts
    """
    client = get_client()
    sheet_id = extract_sheet_id(url_or_id)

    spreadsheet = client.open_by_key(sheet_id)

    worksheets = []
    for ws in spreadsheet.worksheets():
        worksheets.append({
            'title': ws.title,
            'id': ws.id,
            'row_count': ws.row_count,
            'col_count': ws.col_count
        })

    return {
        'id': sheet_id,
        'title': spreadsheet.title,
        'url': spreadsheet.url,
        'worksheets': worksheets
    }


def import_from_sheets(
    url_or_id: str,
    worksheet_name: Optional[str] = None,
    output_path: Optional[str] = None,
    uploads_dir: Optional[str] = None
) -> Tuple[pd.DataFrame, str]:
    """
    Import data from a Google Sheet to a pandas DataFrame.

    Args:
        url_or_id: Google Sheet URL or ID
        worksheet_name: Name of worksheet to import (default: first sheet)
        output_path: Optional path to save as Excel file
        uploads_dir: Optional directory for uploads (used if output_path not provided)

    Returns:
        Tuple of (DataFrame, output_path)
    """
    client = get_client()
    sheet_id = extract_sheet_id(url_or_id)

    spreadsheet = client.open_by_key(sheet_id)

    # Select worksheet
    if worksheet_name:
        worksheet = spreadsheet.worksheet(worksheet_name)
    else:
        worksheet = spreadsheet.sheet1

    # Get all values
    data = worksheet.get_all_records()

    # Convert to DataFrame
    df = pd.DataFrame(data)

    # Generate output path if not provided
    if not output_path:
        if not uploads_dir:
            uploads_dir = os.environ.get('UPLOAD_FOLDER', 'uploads')
        Path(uploads_dir).mkdir(parents=True, exist_ok=True)

        # Use sheet title as filename
        safe_title = re.sub(r'[^\w\-_]', '_', spreadsheet.title)
        output_path = os.path.join(uploads_dir, f"{safe_title}.xlsx")

    # Save as Excel
    df.to_excel(output_path, index=False, sheet_name=worksheet.title)

    return df, output_path


def import_all_worksheets(
    url_or_id: str,
    output_path: Optional[str] = None,
    uploads_dir: Optional[str] = None
) -> Tuple[Dict[str, pd.DataFrame], str, int]:
    """
    Import ALL worksheets from a Google Sheet into a single Excel file.

    Args:
        url_or_id: Google Sheet URL or ID
        output_path: Optional path to save as Excel file
        uploads_dir: Optional directory for uploads (used if output_path not provided)

    Returns:
        Tuple of (dict of DataFrames by sheet name, output_path, total_rows)
    """
    client = get_client()
    sheet_id = extract_sheet_id(url_or_id)

    spreadsheet = client.open_by_key(sheet_id)

    # Generate output path if not provided
    if not output_path:
        if not uploads_dir:
            uploads_dir = os.environ.get('UPLOAD_FOLDER', 'uploads')
        Path(uploads_dir).mkdir(parents=True, exist_ok=True)

        # Use sheet title as filename
        safe_title = re.sub(r'[^\w\-_]', '_', spreadsheet.title)
        output_path = os.path.join(uploads_dir, f"{safe_title}.xlsx")

    # Import all worksheets
    dataframes = {}
    total_rows = 0

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for worksheet in spreadsheet.worksheets():
            try:
                # Get all values from this worksheet
                data = worksheet.get_all_records()
                if data:  # Only add sheets with data
                    df = pd.DataFrame(data)
                    dataframes[worksheet.title] = df
                    total_rows += len(df)
                    df.to_excel(writer, index=False, sheet_name=worksheet.title)
            except Exception as e:
                print(f"Warning: Could not import worksheet '{worksheet.title}': {e}")
                continue

    return dataframes, output_path, total_rows


def export_to_sheets(
    df: pd.DataFrame,
    url_or_id: str,
    worksheet_name: Optional[str] = None,
    create_if_missing: bool = True
) -> Dict:
    """
    Export a pandas DataFrame to a Google Sheet.

    Args:
        df: DataFrame to export
        url_or_id: Google Sheet URL or ID
        worksheet_name: Name of worksheet to export to
        create_if_missing: Create worksheet if it doesn't exist

    Returns:
        dict with export status and info
    """
    client = get_client()
    sheet_id = extract_sheet_id(url_or_id)

    spreadsheet = client.open_by_key(sheet_id)

    # Select or create worksheet
    if worksheet_name:
        try:
            worksheet = spreadsheet.worksheet(worksheet_name)
        except gspread.WorksheetNotFound:
            if create_if_missing:
                worksheet = spreadsheet.add_worksheet(
                    title=worksheet_name,
                    rows=len(df) + 1,
                    cols=len(df.columns)
                )
            else:
                raise
    else:
        worksheet = spreadsheet.sheet1

    # Clear existing data
    worksheet.clear()

    # Prepare data (convert to list of lists)
    headers = df.columns.tolist()
    values = df.fillna('').astype(str).values.tolist()

    # Combine headers and values
    all_data = [headers] + values

    # Update sheet
    worksheet.update(all_data, value_input_option='RAW')

    return {
        'success': True,
        'spreadsheet_id': sheet_id,
        'spreadsheet_title': spreadsheet.title,
        'worksheet': worksheet.title,
        'rows_exported': len(df),
        'url': spreadsheet.url
    }


def sync_excel_to_sheets(
    excel_path: str,
    url_or_id: str,
    sheet_name: Optional[str] = None,
    worksheet_name: Optional[str] = None
) -> Dict:
    """
    Sync an Excel file to a Google Sheet.

    Args:
        excel_path: Path to Excel file
        url_or_id: Google Sheet URL or ID
        sheet_name: Name of sheet in Excel file (default: first sheet)
        worksheet_name: Name of worksheet in Google Sheets

    Returns:
        dict with sync status
    """
    # Read Excel file
    df = pd.read_excel(excel_path, sheet_name=sheet_name or 0)

    # Export to Google Sheets
    result = export_to_sheets(df, url_or_id, worksheet_name)

    result['source_file'] = excel_path
    return result


def get_sheet_columns(url_or_id: str, worksheet_name: Optional[str] = None) -> List[str]:
    """
    Get column headers from a Google Sheet.

    Args:
        url_or_id: Google Sheet URL or ID
        worksheet_name: Name of worksheet (default: first sheet)

    Returns:
        List of column names
    """
    client = get_client()
    sheet_id = extract_sheet_id(url_or_id)

    spreadsheet = client.open_by_key(sheet_id)

    if worksheet_name:
        worksheet = spreadsheet.worksheet(worksheet_name)
    else:
        worksheet = spreadsheet.sheet1

    # Get first row (headers)
    headers = worksheet.row_values(1)
    return headers


def test_connection() -> Dict:
    """
    Test Google Sheets connection.

    Returns:
        dict with connection status
    """
    try:
        client = get_client()
        # Try to list spreadsheets (limited operation)
        return {
            'success': True,
            'message': 'Google Sheets connection successful'
        }
    except FileNotFoundError as e:
        return {
            'success': False,
            'message': f'Service account file not found: {str(e)}'
        }
    except Exception as e:
        return {
            'success': False,
            'message': f'Connection failed: {str(e)}'
        }
