from flask import Flask, render_template, request, jsonify, redirect, url_for
import pandas as pd, math, os, difflib, re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from regenerate import TextRegenerator
from openpyxl.utils import get_column_letter
from urllib.parse import urlencode
import uuid
import yaml
from pathlib import Path

app = Flask(__name__)

# --- Configuration Loading ---
config = {} # Global config dictionary

def load_config(config_path: str = 'config_flash.yaml'):
    # Loads configuration from a YAML file.
    global config
    path = Path(config_path)
    if not path.exists():
        print(f"Warning: Configuration file '{config_path}' not found.")
        config = {} # Ensure config is empty if file not found
        return
    try:
        with path.open('r') as f:
            config = yaml.safe_load(f)
        if not config: # Handle empty config file
            print(f"Warning: Configuration file '{config_path}' is empty.")
            config = {}
    except yaml.YAMLError as e:
        print(f"Error parsing configuration file '{config_path}': {e}")
        config = {} # Reset config on error
    except Exception as e:
        print(f"Error reading configuration file '{config_path}': {e}")
        config = {} # Reset config on error

def get_input_file_path() -> str:
    # Gets the input file path from the loaded configuration.
    default_path = 'input.xlsx' # Keep the original default
    if not config:
        # Don't print warning here, let calling functions handle non-existence
        return default_path
    try:
        path = config.get('file_settings', {}).get('input_file')
        if not path or not isinstance(path, str):
             # Don't print warning here, let calling functions handle non-existence
             return default_path
        return path
    except Exception as e:
        print(f"Error accessing 'input_file' from configuration: {e}. Using default.")
        return default_path

load_config() # Load config when the app starts
# --- End Configuration Loading ---

def compare_text(text1, text2):
    if pd.isna(text1) and pd.isna(text2): return "", "", "same"
    elif pd.isna(text1): 
        replaced_text = str(text2).replace("\n", "<br>")
        return "", f'<span class="added">{replaced_text}</span>', "different"
    elif pd.isna(text2): 
        replaced_text = str(text1).replace("\n", "<br>")
        return f'<span class="removed">{replaced_text}</span>', "", "different"
    
    text1, text2 = str(text1), str(text2)
    if text1 == text2: return text1.replace("\n", "<br>"), text2.replace("\n", "<br>"), "same"
    
    line_break_marker = " ¶ "
    text1_prep = text1.replace('\r\n', '\n').replace('\r', '\n').replace('\n', line_break_marker)
    text2_prep = text2.replace('\r\n', '\n').replace('\r', '\n').replace('\n', line_break_marker)
    
    words1 = re.split(r'(\s+)', text1_prep)
    words2 = re.split(r'(\s+)', text2_prep)
    words1 = [word for word in words1 if word]
    words2 = [word for word in words2 if word]
    
    matcher = difflib.SequenceMatcher(None, words1, words2, autojunk=False)
    result1, result2 = [], []
    diff_id_counter = 0
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        words1_segment = "".join(words1[i1:i2])
        words2_segment = "".join(words2[j1:j2])
        
        if tag == 'equal':
            result1.append(words1_segment)
            result2.append(words2_segment)
        else:
            diff_id = f"diff-{uuid.uuid4()}"
            if words1_segment:
                result1.append(f'<span class="removed" data-diff-id="{diff_id}">{words1_segment}</span>')
            if words2_segment:
                result2.append(f'<span class="added" data-diff-id="{diff_id}">{words2_segment}</span>')
    
    final_text1 = "".join(result1).replace(line_break_marker.strip(), "<br>")
    final_text2 = "".join(result2).replace(line_break_marker.strip(), "<br>")
    
    final_text1 = re.sub(r'<span class="(?:added|removed)" data-diff-id="[^"]*"></span>', '', final_text1)
    final_text2 = re.sub(r'<span class="(?:added|removed)" data-diff-id="[^"]*"></span>', '', final_text2)
    
    return final_text1, final_text2, "different"

def get_cell_color_status():
    input_file = get_input_file_path() # Get path from config
    if not os.path.exists(input_file): return {}
    
    try: wb = load_workbook(input_file)
    except Exception:
        try: wb = load_workbook(input_file, read_only=True)
        except Exception: return {}
    
    if 'hadith' not in wb.sheetnames: return {}
    ws = wb['hadith']
    
    hadith_details_col_idx = analysis3_col_idx = None
    header_row = next(ws.rows)
    for idx, cell in enumerate(header_row):
        col_name = cell.value
        if col_name == 'hadith_details': hadith_details_col_idx = idx
        elif col_name == 'analysis-3': analysis3_col_idx = idx
    
    hadith_details_col_idx = 0 if hadith_details_col_idx is None else hadith_details_col_idx
    analysis3_col_idx = 1 if analysis3_col_idx is None else analysis3_col_idx
    
    color_status = {}
    
    def check_cell_color(cell, row_dict, col_key):
        if not (hasattr(cell, 'fill') and cell.fill and cell.fill.fill_type != 'none'): return
        if not (hasattr(cell.fill.start_color, 'rgb') and cell.fill.start_color.rgb): return
        
        rgb = cell.fill.start_color.rgb
        rgb_str = str(rgb).upper()
        if not (rgb_str and rgb_str != "00000000" and not rgb_str.endswith("000000")): return
        
        row_dict[f'col_{col_key}'] = True
        
        if "FF0000" in rgb_str or "FFFF0000" in rgb_str or rgb_str.endswith("FF0000"): row_dict[f'col_{col_key}_type'] = 'red'
        elif "00FF00" in rgb_str or rgb_str == "FF00FF00": row_dict[f'col_{col_key}_type'] = 'green'
        elif "FFFF00" in rgb_str or rgb_str == "FFFFFF00": row_dict[f'col_{col_key}_type'] = 'yellow'
        elif rgb_str and len(rgb_str) >= 6:
            rgb_part = rgb_str[-6:] if len(rgb_str) > 6 else rgb_str
            r_val, g_val, b_val = rgb_part[0:2], rgb_part[2:4], rgb_part[4:6]
            row_dict[f'col_{col_key}_type'] = 'red' if r_val in ["FF", "F0", "E0"] and g_val in ["00", "10", "20", "30"] and b_val in ["00", "10", "20", "30"] else 'green'

    try:
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            if len(row) > max(hadith_details_col_idx, analysis3_col_idx):
                col_a_cell, col_b_cell = row[hadith_details_col_idx], row[analysis3_col_idx]
                excel_row_idx = row_idx
                color_status[excel_row_idx] = {'col_a': False, 'col_b': False, 'col_a_type': None, 'col_b_type': None}
                check_cell_color(col_a_cell, color_status[excel_row_idx], 'a')
                check_cell_color(col_b_cell, color_status[excel_row_idx], 'b')
    except Exception: pass
    
    return color_status

def get_excel_data(rows_per_page=10, page=1, filter_change_enabled=False, filter_change_value=None, filter_change_lt_value=None, filter_change_from_value=None, filter_change_to_value=None, filter_color_a='any', filter_color_b='any'):
    input_file = get_input_file_path() # Get path from config
    if not os.path.exists(input_file): return [], 0, 0, False

    change_col_exists = False
    try:
        xls = pd.ExcelFile(input_file, engine='openpyxl')
        sheet_name = 'hadith' if 'hadith' in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_name)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        try: df = pd.read_excel(input_file, engine='openpyxl')
        except Exception as e2:
            print(f"Secondary error reading Excel file: {e2}")
            return [], 0, 0, False

    if 'hadith_details' not in df.columns or 'analysis-3' not in df.columns:
        if len(df.columns) >= 2: df = df.rename(columns={df.columns[0]: 'hadith_details', df.columns[1]: 'analysis-3'})
        else: return [], 0, 0, False

    number_col_exists = 'number' in df.columns # Check if 'number' column exists

    if 'change' in df.columns:
        change_col_exists = True
        df['change'] = pd.to_numeric(df['change'], errors='coerce')
    else:
        print("Warning: 'change' column not found in the sheet.")

    if filter_change_enabled and change_col_exists:
        df = df.dropna(subset=['change'])
        if filter_change_value is not None:
            try:
                filter_val = float(filter_change_value)
                df = df[df['change'] > filter_val]
            except (ValueError, TypeError) as e:
                print(f"Invalid filter value for 'change >': {filter_change_value}. Error: {e}")
        if filter_change_lt_value is not None:
            try:
                filter_val = float(filter_change_lt_value)
                df = df[df['change'] < filter_val]
            except (ValueError, TypeError) as e:
                print(f"Invalid filter value for 'change <': {filter_change_lt_value}. Error: {e}")
        if filter_change_from_value is not None and filter_change_to_value is not None:
            try:
                filter_from = float(filter_change_from_value)
                filter_to = float(filter_change_to_value)
                if filter_from <= filter_to:
                    df = df[(df['change'] >= filter_from) & (df['change'] <= filter_to)]
                else: # Handle case where user might swap from/to
                    df = df[(df['change'] >= filter_to) & (df['change'] <= filter_from)]
            except (ValueError, TypeError) as e:
                print(f"Invalid filter values for 'change between': {filter_change_from_value}-{filter_change_to_value}. Error: {e}")
        elif filter_change_from_value is not None: # Only From is specified
             try:
                filter_from = float(filter_change_from_value)
                df = df[df['change'] >= filter_from]
             except (ValueError, TypeError) as e:
                print(f"Invalid filter value for 'change From': {filter_change_from_value}. Error: {e}")
        elif filter_change_to_value is not None: # Only To is specified
             try:
                filter_to = float(filter_change_to_value)
                df = df[df['change'] <= filter_to]
             except (ValueError, TypeError) as e:
                print(f"Invalid filter value for 'change To': {filter_change_to_value}. Error: {e}")

    # Get color status before filtering
    approved_cells = get_cell_color_status()

    # Add color status info directly to the DataFrame for efficient filtering
    df['col_a_approved'] = df.index.map(lambda idx: approved_cells.get(idx + 2, {}).get('col_a', False))
    df['col_a_type'] = df.index.map(lambda idx: approved_cells.get(idx + 2, {}).get('col_a_type', None))
    df['col_b_approved'] = df.index.map(lambda idx: approved_cells.get(idx + 2, {}).get('col_b', False))
    df['col_b_type'] = df.index.map(lambda idx: approved_cells.get(idx + 2, {}).get('col_b_type', None))

    # Apply Color Filters
    if filter_color_a != 'any':
        if filter_color_a == 'none':
            df = df[df['col_a_approved'] == False]
        else: # Specific color
            df = df[(df['col_a_approved'] == True) & (df['col_a_type'] == filter_color_a)]

    if filter_color_b != 'any':
        if filter_color_b == 'none':
            df = df[df['col_b_approved'] == False]
        else: # Specific color
            df = df[(df['col_b_approved'] == True) & (df['col_b_type'] == filter_color_b)]

    df = df.replace('_x000D_', '\n', regex=True).replace(r'\r\n|\r|\n', '\n', regex=True)

    total_rows = len(df)
    total_pages = math.ceil(total_rows / rows_per_page) if rows_per_page > 0 else 1
    page = max(1, min(page, total_pages))

    start_idx = (page - 1) * rows_per_page
    end_idx = start_idx + rows_per_page
    page_data = df.iloc[start_idx:end_idx]

    result = []
    original_indices = df.index[start_idx:end_idx]

    for i, df_idx in enumerate(original_indices):
        row = page_data.iloc[i]
        col_a, col_b = row['hadith_details'], row['analysis-3']
        
        # Get the ID: Use 'number' column if it exists, otherwise fallback to df_idx
        row_id = row['number'] if number_col_exists and 'number' in row and pd.notna(row['number']) else df_idx

        if isinstance(col_a, str): col_a = col_a.replace('_x000D_', '\n').replace('\r\n', '\n').replace('\r', '\n')
        if isinstance(col_b, str): col_b = col_b.replace('_x000D_', '\n').replace('\r\n', '\n').replace('\r', '\n')

        highlighted_a, highlighted_b, status = compare_text(col_a, col_b)
        excel_row_idx = df_idx + 2
        row_approval = approved_cells.get(excel_row_idx, {'col_a': False, 'col_b': False, 'col_a_type': None, 'col_b_type': None})

        result.append({
            'row_idx': df_idx, # Keep original index for internal use (e.g., editing)
            'id': row_id,     # Add the ID to be displayed
            'col_a': col_a if not pd.isna(col_a) else "", 'col_b': col_b if not pd.isna(col_b) else "",
            'highlighted_a': highlighted_a, 'highlighted_b': highlighted_b, 'status': status,
            'col_a_approved': row_approval['col_a'], 'col_b_approved': row_approval['col_b'],
            'col_a_type': row_approval['col_a_type'], 'col_b_type': row_approval['col_b_type']
        })

    return result, total_pages, total_rows, change_col_exists

@app.route('/', methods=['GET'])
def index():
    rows_per_page = request.args.get('rows_per_page', default=10, type=int)
    page = request.args.get('page', default=1, type=int)
    filter_change_enabled = request.args.get('filter_change_enabled') == 'on'
    filter_change_gt_value_str = request.args.get('filter_change_value', default='').strip()
    filter_change_lt_value_str = request.args.get('filter_change_lt_value', default='').strip()
    filter_change_from_value_str = request.args.get('filter_change_from_value', default='').strip()
    filter_change_to_value_str = request.args.get('filter_change_to_value', default='').strip()
    filter_color_a = request.args.get('filter_color_a', default='any').strip().lower()
    filter_color_b = request.args.get('filter_color_b', default='any').strip().lower()

    filter_change_gt_value = None
    filter_change_lt_value = None
    filter_change_from_value = None
    filter_change_to_value = None

    # Validate color filters
    valid_colors = ['any', 'none', 'green', 'red', 'yellow']
    if filter_color_a not in valid_colors: filter_color_a = 'any'
    if filter_color_b not in valid_colors: filter_color_b = 'any'

    if filter_change_enabled:
        if filter_change_gt_value_str:
            try: filter_change_gt_value = float(filter_change_gt_value_str)
            except ValueError: print(f"Warning: Invalid 'change >' filter value: {filter_change_gt_value_str}")
        if filter_change_lt_value_str:
            try: filter_change_lt_value = float(filter_change_lt_value_str)
            except ValueError: print(f"Warning: Invalid 'change <' filter value: {filter_change_lt_value_str}")
        if filter_change_from_value_str:
            try: filter_change_from_value = float(filter_change_from_value_str)
            except ValueError: print(f"Warning: Invalid 'change From' filter value: {filter_change_from_value_str}")
        if filter_change_to_value_str:
            try: filter_change_to_value = float(filter_change_to_value_str)
            except ValueError: print(f"Warning: Invalid 'change To' filter value: {filter_change_to_value_str}")

        if filter_change_gt_value is None and filter_change_lt_value is None and filter_change_from_value is None and filter_change_to_value is None:
             filter_change_enabled = False

    hadith_sheet_missing = False
    input_file = get_input_file_path() # Get path from config
    if os.path.exists(input_file): # Use configured path
        try:
            wb = load_workbook(input_file, read_only=True)
            hadith_sheet_missing = 'hadith' not in wb.sheetnames
            wb.close()
        except Exception: pass

    data, total_pages, total_rows, change_col_exists = get_excel_data(
        rows_per_page,
        page,
        filter_change_enabled,
        filter_change_gt_value,
        filter_change_lt_value,
        filter_change_from_value,
        filter_change_to_value,
        filter_color_a,
        filter_color_b
    )

    query_params = {
        'rows_per_page': rows_per_page,
    }
    if filter_change_enabled:
        query_params['filter_change_enabled'] = 'on'
        if filter_change_gt_value is not None: query_params['filter_change_value'] = filter_change_gt_value_str
        if filter_change_lt_value is not None: query_params['filter_change_lt_value'] = filter_change_lt_value_str
        if filter_change_from_value is not None: query_params['filter_change_from_value'] = filter_change_from_value_str
        if filter_change_to_value is not None: query_params['filter_change_to_value'] = filter_change_to_value_str
    
    # Add color filters to query params if they are not 'any'
    if filter_color_a != 'any': query_params['filter_color_a'] = filter_color_a
    if filter_color_b != 'any': query_params['filter_color_b'] = filter_color_b

    return render_template('index.html',
                          data=data,
                          total_pages=total_pages,
                          current_page=page,
                          rows_per_page=rows_per_page,
                          total_rows=total_rows,
                          hadith_sheet_missing=hadith_sheet_missing,
                          filter_change_enabled=filter_change_enabled,
                          filter_change_value=filter_change_gt_value_str,
                          filter_change_lt_value=filter_change_lt_value_str,
                          filter_change_from_value=filter_change_from_value_str,
                          filter_change_to_value=filter_change_to_value_str,
                          filter_color_a=filter_color_a,
                          filter_color_b=filter_color_b,
                          change_col_exists=change_col_exists,
                          query_params=query_params
                          )

@app.route('/edit', methods=['POST'])
def edit_cell():
    row_idx, new_text = request.form.get('row_idx', type=int), request.form.get('text', '')
    new_text = new_text.replace('<br>', '\n').replace('<br/>', '\n').replace('\r\n', '\n').replace('\r', '\n')
    
    input_file = get_input_file_path() # Get path from config
    if not os.path.exists(input_file): return jsonify({'status': 'error', 'message': 'Excel file not found'})
    
    try:
        wb = load_workbook(input_file)
        if 'hadith' not in wb.sheetnames: return jsonify({'status': 'error', 'message': 'Hadith sheet not found in Excel file'})
        ws = wb['hadith']
        
        header = next(ws.rows)
        analysis3_col_idx, hadith_details_col_idx = 1, 0
        
        for idx, cell in enumerate(header):
            if cell.value == 'analysis-3':
                analysis3_col_idx = idx
                break
        
        excel_row = row_idx + 2
        cell_address = f'{chr(65 + analysis3_col_idx)}{excel_row}'
        ws[cell_address].value = new_text
        wb.save(input_file)
        
        for idx, cell in enumerate(header):
            if cell.value == 'hadith_details':
                hadith_details_col_idx = idx
                break
        
        col_a_cell = ws[f'{chr(65 + hadith_details_col_idx)}{excel_row}']
        col_a_text = str(col_a_cell.value) if col_a_cell.value is not None else ''
        highlighted_a, highlighted_b, status = compare_text(col_a_text, new_text)
        
        return jsonify({'status': 'success', 'highlighted_html': highlighted_b, 'diff_status': status})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/approve', methods=['POST'])
def approve_cell():
    row_idx = request.form.get('row_idx', type=int)
    column = request.form.get('column')
    approval_type = request.form.get('approval_type', 'green')
    
    input_file = get_input_file_path() # Get path from config
    try:
        if not os.path.exists(input_file): return jsonify({'status': 'error', 'message': 'Excel file not found'})
        
        wb = load_workbook(input_file)
        if 'hadith' not in wb.sheetnames: return jsonify({'status': 'error', 'message': 'Hadith sheet not found in Excel file'})
        
        ws = wb['hadith']
        header_row = next(ws.rows)
        hadith_details_col_idx, analysis3_col_idx = 0, 1
        
        for idx, cell in enumerate(header_row):
            col_name = cell.value
            if col_name == 'hadith_details': hadith_details_col_idx = idx
            elif col_name == 'analysis-3': analysis3_col_idx = idx
        
        excel_row = row_idx + 2
        column_idx = hadith_details_col_idx if column == 'a' else analysis3_col_idx
        cell_address = f'{chr(65 + column_idx)}{excel_row}'
        
        colors = {'green': "00FF00", 'yellow': "FFFF00", 'red': "FFFF0000"}
        color = colors.get(approval_type, "00FF00")
        
        ws[cell_address].fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        wb.save(input_file)
        
        return jsonify({'status': 'success', 'message': 'Cell approved successfully', 
                       'row_idx': row_idx, 'column': column, 'approval_type': approval_type})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/reset_cell', methods=['POST'])
def reset_cell():
    row_idx, column = request.form.get('row_idx', type=int), request.form.get('column')
    
    input_file = get_input_file_path() # Get path from config
    if not os.path.exists(input_file): return jsonify({'status': 'error', 'message': 'Excel file not found'})
    
    try:
        wb = load_workbook(input_file)
        if 'hadith' not in wb.sheetnames: return jsonify({'status': 'error', 'message': 'Hadith sheet not found in Excel file'})
        
        ws = wb['hadith']
        header_row = next(ws.rows)
        hadith_details_col_idx, analysis3_col_idx = 0, 1
        
        for idx, cell in enumerate(header_row):
            col_name = cell.value
            if col_name == 'hadith_details': hadith_details_col_idx = idx
            elif col_name == 'analysis-3': analysis3_col_idx = idx
        
        excel_row = row_idx + 2
        column_idx = hadith_details_col_idx if column == 'a' else analysis3_col_idx
        cell_address = f'{chr(65 + column_idx)}{excel_row}'
        
        ws[cell_address].fill = PatternFill(fill_type=None)
        wb.save(input_file)
        
        return jsonify({'status': 'success', 'message': 'Cell color reset successfully', 'row_idx': row_idx, 'column': column})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/save_selection', methods=['POST'])
def save_selection():
    selected_text = request.form.get('selected_text', '')
    
    input_file = get_input_file_path() # Get path from config
    if not selected_text.strip(): return jsonify({'status': 'error', 'message': 'No text selected'})
    if not os.path.exists(input_file): return jsonify({'status': 'error', 'message': 'Excel file not found'})
    
    try:
        wb = load_workbook(input_file)
        
        if 'hadith' not in wb.sheetnames:
            ws_hadith = wb.create_sheet('hadith')
            ws_hadith['A1'], ws_hadith['B1'] = 'hadith_details', 'analysis-3'
        
        if 'words' not in wb.sheetnames:
            ws = wb.create_sheet('words')
            ws['A1'] = 'word_list'
        else:
            ws = wb['words']
            if ws['A1'].value != 'word_list': ws['A1'] = 'word_list'
        
        row = 2
        while ws[f'A{row}'].value: row += 1
            
        ws[f'A{row}'] = selected_text
        wb.save(input_file)
        
        return jsonify({'status': 'success', 'message': 'Text saved successfully', 'text': selected_text, 'row': row})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/preview_diff', methods=['POST'])
def preview_diff():
    text1, text2 = request.form.get('text1', ''), request.form.get('text2', '')
    try:
        highlighted_a, highlighted_b, status = compare_text(text1, text2)
        return jsonify({'highlighted_a': highlighted_a, 'highlighted_b': highlighted_b, 'status': status})
    except Exception as e:
        return jsonify({'status': 'error', 'message': f"Error comparing texts: {str(e)}"})

@app.route('/regenerate_cell', methods=['POST'])
def regenerate_cell():
    row_idx = request.form.get('row_idx', type=int)
    try:
        regenerator = TextRegenerator(config_path='config_flash.yaml')
        new_text, original_text = regenerator.regenerate_text(row_idx)

        input_file = get_input_file_path()
        if not os.path.exists(input_file):
            return jsonify({'status': 'error', 'message': 'Excel file not found after regeneration'})

        wb = load_workbook(input_file)
        if 'hadith' not in wb.sheetnames:
            return jsonify({'status': 'error', 'message': 'Hadith sheet not found'})
        ws = wb['hadith']

        header = next(ws.rows)
        analysis3_col_idx = 1
        hadith_details_col_idx = 0
        for idx, cell in enumerate(header):
            if cell.value == 'analysis-3':
                analysis3_col_idx = idx
            elif cell.value == 'hadith_details':
                hadith_details_col_idx = idx

        excel_row = row_idx + 2
        cell_address = f'{get_column_letter(analysis3_col_idx + 1)}{excel_row}'
        ws[cell_address].value = new_text
        ws[cell_address].fill = PatternFill(fill_type=None)  # Clear existing fill

        wb.save(input_file)

        # Fetch updated color status for Column B
        color_status = get_cell_color_status()
        row_approval = color_status.get(excel_row, {'col_b': False, 'col_b_type': None})
        col_b_approved = row_approval['col_b']
        col_b_type = row_approval['col_b_type']

        # Get original text from Column A for comparison
        col_a_cell = ws[f'{get_column_letter(hadith_details_col_idx + 1)}{excel_row}']
        col_a_text = str(col_a_cell.value) if col_a_cell.value is not None else ''
        highlighted_a, highlighted_b, status = compare_text(col_a_text, new_text)

        return jsonify({
            'status': 'success',
            'new_text': new_text,
            'highlighted_html': highlighted_b,
            'diff_status': status,
            'col_b_approved': col_b_approved,  # Add approval status
            'col_b_type': col_b_type          # Add approval type
        })

    except FileNotFoundError as e:
        return jsonify({'status': 'error', 'message': str(e)})
    except ValueError as e:
        return jsonify({'status': 'error', 'message': str(e)})
    except Exception as e:
        print(f"Error during regeneration or file update for row {row_idx}: {type(e).__name__} - {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'status': 'error', 'message': f'An unexpected error occurred: {str(e)}'})

@app.route('/get_comment', methods=['GET'])
def get_comment():
    row_idx = request.args.get('row_idx', type=int)
    
    input_file = get_input_file_path() # Get path from config
    if not os.path.exists(input_file): # Use configured path
        return jsonify({'status': 'error', 'message': 'Excel file not found'})
    
    try:
        df = pd.read_excel(input_file, sheet_name='hadith')
        
        if 'comments' not in df.columns:
            return jsonify({'comment': '', 'status': 'success'})
        
        comment = df.loc[row_idx, 'comments'] if pd.notna(df.loc[row_idx, 'comments']) else ''
        
        return jsonify({'comment': comment, 'status': 'success'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/save_comment', methods=['POST'])
def save_comment():
    row_idx = request.form.get('row_idx', type=int)
    comment = request.form.get('comment', '')
    
    input_file = get_input_file_path() # Get path from config
    if not os.path.exists(input_file): # Use configured path
        return jsonify({'status': 'error', 'message': 'Excel file not found'})
    
    try:
        wb = load_workbook(input_file)
        ws = wb['hadith'] if 'hadith' in wb.sheetnames else wb.active
        
        header_row = next(ws.rows)
        comments_col_idx = None
        
        for idx, cell in enumerate(header_row):
            if cell.value == 'comments':
                comments_col_idx = idx
                break
        
        if comments_col_idx is None:
            comments_col_idx = len(header_row)
            comments_col_letter = get_column_letter(comments_col_idx + 1)
            ws[f'{comments_col_letter}1'] = 'comments'
        
        excel_row = row_idx + 2
        comments_col_letter = get_column_letter(comments_col_idx + 1)
        ws[f'{comments_col_letter}{excel_row}'] = comment
        
        wb.save(input_file)
        
        return jsonify({'status': 'success', 'message': 'Comment saved successfully'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.context_processor
def utility_processor():
    return dict(urlencode=urlencode)

if __name__ == '__main__': app.run(debug=True,port=8000)
