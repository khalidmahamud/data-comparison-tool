from flask import Flask, render_template, request, jsonify, redirect, url_for
import pandas as pd, math, os, difflib, re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from urllib.parse import urlencode
import uuid

from pathlib import Path
from src.prompt import inject_variables
from src.ai import ask
from src.generate_cell import generate, extract_standard_letters
from src.config import config, load_config

app = Flask(__name__)

# --- Configuration Loading ---
# Global variable to track the currently selected chunk
current_chunk = None

# Function to get all available chunks
def get_available_chunks():
    chunks_dir = 'chunks'
    if not os.path.exists(chunks_dir):
        return []
    
    chunks = []
    chunk_pattern = re.compile(r'chunk_(\d+)_rows_(\d+)-(\d+)\.xlsx')
    
    for filename in os.listdir(chunks_dir):
        if chunk_pattern.match(filename):
            chunk_match = chunk_pattern.match(filename)
            chunk_num = int(chunk_match.group(1))
            start_row = int(chunk_match.group(2))
            end_row = int(chunk_match.group(3))
            
            chunks.append({
                'filename': os.path.join(chunks_dir, filename),
                'chunk_num': chunk_num,
                'display_name': f"C{chunk_num} ({start_row}-{end_row})",
                'start_row': start_row,
                'end_row': end_row
            })
    
    # Sort by chunk number
    chunks.sort(key=lambda x: x['chunk_num'])
    return chunks

# Initialize current_chunk to the first available chunk
chunks = get_available_chunks()
if chunks:
    current_chunk = chunks[0]['filename']

def reload_config(config_path: str = 'config_flash.yaml'):
    # Reloads the configuration from a YAML file.
    global config
    try:
        config = load_config(config_path)
    except Exception as e:
        print(f"Error loading configuration: {e}")

def get_input_file_path() -> str:
    global current_chunk
    if current_chunk:
        return current_chunk

    # Gets the input file path from the loaded configuration.
    default_path = 'input.xlsx' # Keep the original default
    try:
        return config.file_settings.input_file
    except Exception as e:
        print(f"Error accessing 'input_file' from configuration: {e}. Using default.")
        return default_path

def get_sheet_name() -> str:
    # Get the sheet name from configuration
    try:
        return config.excel_settings.sheet_name
    except Exception as e:
        print(f"Error accessing sheet name from configuration: {e}. Using default.")
        return 'hadith'

def get_column_name(column_key: str) -> str:
    # Get the column name for a specific key from configuration
    try:
        return config.excel_settings.columns.get(column_key, column_key)
    except Exception as e:
        # Default fallbacks for essential columns
        defaults = {
            'primary_text': 'hadith_details',
            'secondary_text': 'analysis-3',
            'ratio': 'ratio',
            'number': 'number',
            'arabic_text': 'arabic_text'
        }
        print(f"Error accessing column name from configuration: {e}. Using default.")
        return defaults.get(column_key, column_key)

# Ensure configuration is loaded when the app starts
reload_config()
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
    
    sheet_name = get_sheet_name()
    if sheet_name not in wb.sheetnames: return {}
    ws = wb[sheet_name]
    
    primary_text_col_name = get_column_name('primary_text')
    secondary_text_col_name = get_column_name('secondary_text')
    
    primary_text_col_idx = secondary_text_col_idx = None
    header_row = next(ws.rows)
    for idx, cell in enumerate(header_row):
        col_name = cell.value
        if col_name == primary_text_col_name: primary_text_col_idx = idx
        elif col_name == secondary_text_col_name: secondary_text_col_idx = idx
    
    primary_text_col_idx = 0 if primary_text_col_idx is None else primary_text_col_idx
    secondary_text_col_idx = 1 if secondary_text_col_idx is None else secondary_text_col_idx
    
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
            if len(row) > max(primary_text_col_idx, secondary_text_col_idx):
                col_a_cell, col_b_cell = row[primary_text_col_idx], row[secondary_text_col_idx]
                excel_row_idx = row_idx
                color_status[excel_row_idx] = {'col_a': False, 'col_b': False, 'col_a_type': None, 'col_b_type': None}
                check_cell_color(col_a_cell, color_status[excel_row_idx], 'a')
                check_cell_color(col_b_cell, color_status[excel_row_idx], 'b')
    except Exception: pass
    
    return color_status

def get_excel_data(rows_per_page=10, page=1, filter_change_enabled=False, filter_change_value=None, filter_change_lt_value=None, filter_change_from_value=None, filter_change_to_value=None, filter_color_a='any', filter_color_b='any', sort_order='asc', filter_id=None):
    input_file = get_input_file_path() # Get path from config
    if not os.path.exists(input_file): return [], 0, 0, False

    change_col_exists = False
    try:
        xls = pd.ExcelFile(input_file, engine='openpyxl')
        sheet_name = get_sheet_name()
        sheet_name = sheet_name if sheet_name in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_name)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        try: df = pd.read_excel(input_file, engine='openpyxl')
        except Exception as e2:
            print(f"Secondary error reading Excel file: {e2}")
            return [], 0, 0, False

    primary_text_col = get_column_name('primary_text')
    secondary_text_col = get_column_name('secondary_text')
    ratio_col = get_column_name('ratio')
    number_col = get_column_name('number')

    if primary_text_col not in df.columns or secondary_text_col not in df.columns:
        if len(df.columns) >= 2: df = df.rename(columns={df.columns[0]: primary_text_col, df.columns[1]: secondary_text_col})
        else: return [], 0, 0, False

    if ratio_col not in df.columns:
        df[ratio_col] = df.apply(lambda row: difflib.SequenceMatcher(
            None, 
            str(row[primary_text_col]) if pd.notna(row[primary_text_col]) else "",
            str(row[secondary_text_col]) if pd.notna(row[secondary_text_col]) else "",
            autojunk=False
        ).ratio() * 100, axis=1)

        
        
        # Save the ratio column back to Excel file while preserving formatting
        try:
            wb = load_workbook(input_file)
            sheet_name = get_sheet_name()
            if sheet_name not in wb.sheetnames:
                print(f"Warning: '{sheet_name}' sheet not found in Excel file")
            else:
                ws = wb[sheet_name]
                
                # Find the last column index
                last_col_idx = len(next(ws.rows))
                ratio_col_letter = get_column_letter(last_col_idx + 1)
                
                # Add ratio header
                ws[f'{ratio_col_letter}1'] = ratio_col
                
                # Add ratio values for each row
                for idx, ratio in enumerate(df[ratio_col], start=2):
                    ws[f'{ratio_col_letter}{idx}'] = ratio
                
                wb.save(input_file)
        except Exception as e:
            print(f"Warning: Could not save ratio column to Excel file: {e}")

    number_col_exists = number_col in df.columns # Check if 'number' column exists

    if 'change' in df.columns:
        change_col_exists = True
        df['change'] = pd.to_numeric(df['change'], errors='coerce')

    # Sort by ratio based on sort_order parameter
    if sort_order == 'asc':
        df = df.sort_values(by=ratio_col, ascending=True)
    elif sort_order == 'desc':
        df = df.sort_values(by=ratio_col, ascending=False)
    # If sort_order is 'none', no sorting is applied - retain original order

    if filter_change_enabled:
        df = df.dropna(subset=[ratio_col])
        if filter_change_value is not None:
            try:
                filter_val = float(filter_change_value)
                df = df[df[ratio_col] > filter_val]
            except (ValueError, TypeError) as e:
                print(f"Invalid filter value for 'change >': {filter_change_value}. Error: {e}")
        if filter_change_lt_value is not None:
            try:
                filter_val = float(filter_change_lt_value)
                df = df[df[ratio_col] < filter_val]
            except (ValueError, TypeError) as e:
                print(f"Invalid filter value for 'change <': {filter_change_lt_value}. Error: {e}")
        if filter_change_from_value is not None and filter_change_to_value is not None:
            try:
                filter_from = float(filter_change_from_value)
                filter_to = float(filter_change_to_value)
                if filter_from <= filter_to:
                    df = df[(df[ratio_col] >= filter_from) & (df[ratio_col] <= filter_to)]
                else: # Handle case where user might swap from/to
                    df = df[(df[ratio_col] >= filter_to) & (df[ratio_col] <= filter_from)]
            except (ValueError, TypeError) as e:
                print(f"Invalid filter values for 'change between': {filter_change_from_value}-{filter_change_to_value}. Error: {e}")
        elif filter_change_from_value is not None: # Only From is specified
             try:
                filter_from = float(filter_change_from_value)
                df = df[df[ratio_col] >= filter_from]
             except (ValueError, TypeError) as e:
                print(f"Invalid filter value for 'change From': {filter_change_from_value}. Error: {e}")
        elif filter_change_to_value is not None: # Only To is specified
             try:
                filter_to = float(filter_change_to_value)
                df = df[df[ratio_col] <= filter_to]
             except (ValueError, TypeError) as e:
                print(f"Invalid filter value for 'change To': {filter_change_to_value}. Error: {e}")

    # Get color status before filtering
    approved_cells = get_cell_color_status()

    # Apply ID filter if provided
    if filter_id is not None and filter_id != "":
        # First try to filter by 'number' column if it exists
        if number_col_exists:
            # Convert filter_id to the same type as in the DataFrame for comparison
            sample_type = type(df[number_col].iloc[0]) if not df.empty and not pd.isna(df[number_col].iloc[0]) else None
            if sample_type == int:
                try:
                    filter_id_value = int(filter_id)
                    df = df[df[number_col] == filter_id_value]
                except (ValueError, TypeError):
                    # If conversion fails, try exact string match
                    df = df[df[number_col].astype(str) == str(filter_id)]
            elif sample_type == float:
                try:
                    filter_id_value = float(filter_id)
                    df = df[df[number_col] == filter_id_value]
                except (ValueError, TypeError):
                    # If conversion fails, try exact string match
                    df = df[df[number_col].astype(str) == str(filter_id)]
            else:
                # For any other type, use string comparison
                df = df[df[number_col].astype(str) == str(filter_id)]
        
        # If number column doesn't exist or no match was found, try to filter by index
        if len(df) == 0 or not number_col_exists:
            try:
                # Try to convert filter_id to integer for index filtering
                filter_idx = int(filter_id)
                if filter_idx in df.index:
                    df = df.loc[[filter_idx]]
            except (ValueError, TypeError):
                # If filter_id is not a valid integer, no rows will match
                if not number_col_exists:  # Only apply empty filter if we haven't found matches already
                    df = df.head(0)  # Empty DataFrame with same structure

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
        col_a, col_b = row[primary_text_col], row[secondary_text_col]
        
        # Get the ID: Use 'number' column if it exists, otherwise fallback to df_idx
        row_id = row[number_col] if number_col_exists and number_col in row and pd.notna(row[number_col]) else df_idx

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
            'col_a_type': row_approval['col_a_type'], 'col_b_type': row_approval['col_b_type'],
            'ratio': row[ratio_col] if ratio_col in row else None
        })

    return result, total_pages, total_rows, change_col_exists

@app.route('/', methods=['GET'])
def index():
    rows_per_page = request.args.get('rows_per_page', default=10, type=int)
    page = request.args.get('page', default=1, type=int)
    filter_change_enabled = request.args.get('filter_change_enabled') == 'on'
    filter_change_gt_value_str = request.args.get('filter_change_gt_value', default='').strip()
    filter_change_lt_value_str = request.args.get('filter_change_lt_value', default='').strip()
    filter_change_from_value_str = request.args.get('filter_change_from_value', default='').strip()
    filter_change_to_value_str = request.args.get('filter_change_to_value', default='').strip()
    filter_color_a = request.args.get('filter_color_a', default='any').strip().lower()
    filter_color_b = request.args.get('filter_color_b', default='any').strip().lower()
    sort_order = request.args.get('sort_order', default='asc').strip().lower()
    filter_id = request.args.get('filter_id', default=None)
    
    # If filter_id is provided but empty, set it to None
    if filter_id and filter_id.strip() == "":
        filter_id = None

    filter_change_gt_value = None
    filter_change_lt_value = None
    filter_change_from_value = None
    filter_change_to_value = None

    # Validate color filters
    valid_colors = ['any', 'none', 'green', 'red', 'yellow']
    if filter_color_a not in valid_colors: filter_color_a = 'any'
    if filter_color_b not in valid_colors: filter_color_b = 'any'
    
    # Validate sort_order
    valid_sort_orders = ['asc', 'desc', 'none']
    if sort_order not in valid_sort_orders: sort_order = 'asc'

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

    data_sheet_missing = False
    input_file = get_input_file_path() # Get path from config
    if os.path.exists(input_file): # Use configured path
        try:
            wb = load_workbook(input_file, read_only=True)
            sheet_name = get_sheet_name()
            data_sheet_missing = sheet_name not in wb.sheetnames
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
        filter_color_b,
        sort_order,
        filter_id
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
    
    # Add sort_order to query params if it's not the default 'asc'
    if sort_order != 'asc': query_params['sort_order'] = sort_order

    # Add ID filter to query params if it's not None
    if filter_id is not None:
        query_params['filter_id'] = filter_id

    return render_template('index.html',
                          data=data,
                          total_pages=total_pages,
                          current_page=page,
                          rows_per_page=rows_per_page,
                          total_rows=total_rows,
                          hadith_sheet_missing=data_sheet_missing,
                          filter_change_enabled=filter_change_enabled,
                          filter_change_gt_value=filter_change_gt_value_str,
                          filter_change_lt_value=filter_change_lt_value_str,
                          filter_change_from_value=filter_change_from_value_str,
                          filter_change_to_value=filter_change_to_value_str,
                          filter_color_a=filter_color_a,
                          filter_color_b=filter_color_b,
                          filter_id=filter_id,
                          sort_order=sort_order,
                          change_col_exists=change_col_exists,
                          query_params=query_params,
                          available_chunks=get_available_chunks(),
                          current_chunk=current_chunk
                          )

@app.route('/edit', methods=['POST'])
def edit_cell():
    row_idx, new_text = request.form.get('row_idx', type=int), request.form.get('text', '')
    new_text = new_text.replace('<br>', '\n').replace('<br/>', '\n').replace('\r\n', '\n').replace('\r', '\n')
    
    input_file = get_input_file_path() # Get path from config
    if not os.path.exists(input_file): return jsonify({'status': 'error', 'message': 'Excel file not found'})
    
    try:
        wb = load_workbook(input_file)
        sheet_name = get_sheet_name()
        if sheet_name not in wb.sheetnames: return jsonify({'status': 'error', 'message': f'{sheet_name} sheet not found in Excel file'})
        ws = wb[sheet_name]
        
        header = next(ws.rows)
        secondary_text_col_idx, primary_text_col_idx = 1, 0
        
        primary_text_col_name = get_column_name('primary_text')
        secondary_text_col_name = get_column_name('secondary_text')
        
        for idx, cell in enumerate(header):
            if cell.value == secondary_text_col_name:
                secondary_text_col_idx = idx
                break
        
        for idx, cell in enumerate(header):
            if cell.value == primary_text_col_name:
                primary_text_col_idx = idx
                break
        
        excel_row = row_idx + 2
        cell_address = f'{chr(65 + secondary_text_col_idx)}{excel_row}'
        ws[cell_address].value = new_text
        wb.save(input_file)
        
        col_a_cell = ws[f'{chr(65 + primary_text_col_idx)}{excel_row}']
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
        sheet_name = get_sheet_name()
        if sheet_name not in wb.sheetnames: return jsonify({'status': 'error', 'message': f'{sheet_name} sheet not found in Excel file'})
        
        ws = wb[sheet_name]
        header_row = next(ws.rows)
        
        primary_text_col_name = get_column_name('primary_text')
        secondary_text_col_name = get_column_name('secondary_text')
        
        primary_text_col_idx, secondary_text_col_idx = 0, 1
        
        for idx, cell in enumerate(header_row):
            col_name = cell.value
            if col_name == primary_text_col_name: primary_text_col_idx = idx
            elif col_name == secondary_text_col_name: secondary_text_col_idx = idx
        
        excel_row = row_idx + 2
        column_idx = primary_text_col_idx if column == 'a' else secondary_text_col_idx
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
        sheet_name = get_sheet_name()
        if sheet_name not in wb.sheetnames: return jsonify({'status': 'error', 'message': f'{sheet_name} sheet not found in Excel file'})
        
        ws = wb[sheet_name]
        header_row = next(ws.rows)
        
        primary_text_col_name = get_column_name('primary_text')
        secondary_text_col_name = get_column_name('secondary_text')
        
        primary_text_col_idx, secondary_text_col_idx = 0, 1
        
        for idx, cell in enumerate(header_row):
            col_name = cell.value
            if col_name == primary_text_col_name: primary_text_col_idx = idx
            elif col_name == secondary_text_col_name: secondary_text_col_idx = idx
        
        excel_row = row_idx + 2
        column_idx = primary_text_col_idx if column == 'a' else secondary_text_col_idx
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
        sheet_name = get_sheet_name()
        
        if sheet_name not in wb.sheetnames:
            ws_data = wb.create_sheet(sheet_name)
            primary_text_col = get_column_name('primary_text')
            secondary_text_col = get_column_name('secondary_text')
            ws_data['A1'], ws_data['B1'] = primary_text_col, secondary_text_col
        
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
        new_text = generate(row_idx, get_input_file_path()).strip()

        input_file = get_input_file_path()
        
        if not os.path.exists(input_file):
            return jsonify({'status': 'error', 'message': 'Excel file not found after regeneration'})

        wb = load_workbook(input_file)
        sheet_name = get_sheet_name()
        if sheet_name not in wb.sheetnames:
            return jsonify({'status': 'error', 'message': f'{sheet_name} sheet not found in Excel file'})
        ws = wb[sheet_name]

        header = next(ws.rows)
        primary_text_col_name = get_column_name('primary_text')
        secondary_text_col_name = get_column_name('secondary_text')
        
        secondary_text_col_idx = 1
        primary_text_col_idx = 0
        
        for idx, cell in enumerate(header):
            if cell.value == secondary_text_col_name:
                secondary_text_col_idx = idx
            elif cell.value == primary_text_col_name:
                primary_text_col_idx = idx

        excel_row = row_idx + 2
        cell_address = f'{get_column_letter(secondary_text_col_idx + 1)}{excel_row}'
        ws[cell_address].value = new_text
        ws[cell_address].fill = PatternFill(fill_type=None)  # Clear existing fill

        wb.save(input_file)

        # Fetch updated color status for Column B
        color_status = get_cell_color_status()
        row_approval = color_status.get(excel_row, {'col_b': False, 'col_b_type': None})
        col_b_approved = row_approval['col_b']
        col_b_type = row_approval['col_b_type']

        # Get original text from Column A for comparison
        col_a_cell = ws[f'{get_column_letter(primary_text_col_idx + 1)}{excel_row}']
        col_a_text = str(col_a_cell.value) if col_a_cell.value is not None else ''
        highlighted_a, highlighted_b, status = compare_text(col_a_text, new_text)

        return jsonify({
            'status': 'success',
            'new_text': new_text,
            'highlighted_html': highlighted_b,
            'highlighted_a_html': highlighted_a,  # Include highlighted HTML for column A
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

@app.route('/regenerate_multiple_cells', methods=['POST'])
def regenerate_multiple_cells():
    row_ids = request.json.get('row_ids', [])

    print(f"Regenerating rows: {row_ids}")
    
    if not row_ids:
        return jsonify({'status': 'error', 'message': 'No row IDs provided'})
    
    try:
        import concurrent.futures
        
        input_file = get_input_file_path()
        results = []
        
        if not os.path.exists(input_file):
            return jsonify({'status': 'error', 'message': 'Excel file not found'})
        
        # First, generate all the new texts in parallel
        generated_texts = {}
        
        def generate_text_for_row(row_idx):
            try:
                print(f"Generating text for row: {row_idx}")
                new_text = generate(row_idx, input_file).strip()
                return {'status': 'success', 'row_idx': row_idx, 'new_text': new_text}
            except Exception as e:
                import traceback
                return {
                    'status': 'error',
                    'row_idx': row_idx,
                    'message': str(e),
                    'traceback': traceback.format_exc()
                }
        
        # Generate all texts in parallel
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(10, len(row_ids))) as executor:
            # Submit all generation tasks
            future_to_row = {executor.submit(generate_text_for_row, row_idx): row_idx for row_idx in row_ids}
            
            # Collect results as they complete
            for future in concurrent.futures.as_completed(future_to_row):
                row_idx = future_to_row[future]
                try:
                    result = future.result()
                    if result['status'] == 'success':
                        generated_texts[row_idx] = result['new_text']
                    else:
                        results.append(result)  # Store error results
                except Exception as e:
                    results.append({
                        'status': 'error',
                        'row_idx': row_idx,
                        'message': str(e)
                    })
        
        # If there are successful generations, update the Excel file only once
        if generated_texts:
            try:
                # Load workbook once
                wb = load_workbook(input_file)
                sheet_name = get_sheet_name()
                
                if sheet_name not in wb.sheetnames:
                    return jsonify({'status': 'error', 'message': f'{sheet_name} sheet not found'})
                
                ws = wb[sheet_name]
                
                # Find column indices
                header = next(ws.rows)
                primary_text_col_name = get_column_name('primary_text')
                secondary_text_col_name = get_column_name('secondary_text')
                
                secondary_text_col_idx = 1
                primary_text_col_idx = 0
                
                for idx, cell in enumerate(header):
                    if cell.value == secondary_text_col_name:
                        secondary_text_col_idx = idx
                    elif cell.value == primary_text_col_name:
                        primary_text_col_idx = idx
                
                # Update all cells at once
                for row_idx, new_text in generated_texts.items():
                    excel_row = row_idx + 2
                    cell_address = f'{get_column_letter(secondary_text_col_idx + 1)}{excel_row}'
                    ws[cell_address].value = new_text
                    ws[cell_address].fill = PatternFill(fill_type=None)  # Clear existing fill
                
                # Save the workbook once after all updates
                wb.save(input_file)
                
                # Now collect all comparison results
                for row_idx, new_text in generated_texts.items():
                    excel_row = row_idx + 2
                    
                    # Get original text for comparison
                    col_a_cell = ws[f'{get_column_letter(primary_text_col_idx + 1)}{excel_row}']
                    col_a_text = str(col_a_cell.value) if col_a_cell.value is not None else ''
                    highlighted_a, highlighted_b, status = compare_text(col_a_text, new_text)
                    
                    # Fetch color status for this row
                    color_status = get_cell_color_status()
                    row_approval = color_status.get(excel_row, {'col_b': False, 'col_b_type': None})
                    col_b_approved = row_approval['col_b']
                    col_b_type = row_approval['col_b_type']
                    
                    results.append({
                        'status': 'success',
                        'row_idx': row_idx,
                        'new_text': new_text,
                        'highlighted_html': highlighted_b,
                        'highlighted_a_html': highlighted_a,
                        'diff_status': status,
                        'col_b_approved': col_b_approved,
                        'col_b_type': col_b_type
                    })
                
            except Exception as e:
                import traceback
                error_message = f"Error updating Excel file: {str(e)}"
                traceback.print_exc()
                
                # Add error for each row that was not already recorded as an error
                for row_idx in generated_texts.keys():
                    if not any(r.get('row_idx') == row_idx and r.get('status') == 'error' for r in results):
                        results.append({
                            'status': 'error',
                            'row_idx': row_idx,
                            'message': error_message
                        })
        
        success_count = sum(1 for r in results if r['status'] == 'success')
        
        return jsonify({
            'status': 'success',
            'message': f'Successfully processed {success_count} of {len(row_ids)} rows',
            'results': results
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({
            'status': 'error',
            'message': f'An unexpected error occurred: {str(e)}'
        })

@app.route('/get_comment', methods=['GET'])
def get_comment():
    row_idx = request.args.get('row_idx', type=int)
    
    input_file = get_input_file_path() # Get path from config
    if not os.path.exists(input_file): # Use configured path
        return jsonify({'status': 'error', 'message': 'Excel file not found'})
    
    try:
        sheet_name = get_sheet_name()
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        
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
        sheet_name = get_sheet_name()
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        
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
    

@app.route('/get_arabic_text', methods=['GET'])
def get_arabic_text():
    try:
        row_idx = request.args.get('row_idx')
        
        if not row_idx:
            return jsonify({'status': 'error', 'message': 'Row index is required'})
        
        input_file = get_input_file_path()
        if not os.path.exists(input_file):
            return jsonify({'status': 'error', 'message': 'Input file not found'})
        
        try:
            # Read the Excel file
            xls = pd.ExcelFile(input_file, engine='openpyxl')
            sheet_name = get_sheet_name()
            sheet_name = sheet_name if sheet_name in xls.sheet_names else xls.sheet_names[0]
            df = pd.read_excel(xls, sheet_name=sheet_name)
        except Exception as e:
            return jsonify({'status': 'error', 'message': f'Error reading Excel file: {str(e)}'})
        
        # Get the Arabic column name from config
        arabic_column = get_column_name('arabic_text')
        
        # Check for Arabic column in order of preference
        if arabic_column not in df.columns:
            # First check for common Arabic column names if config value not found
            if 'arabic_text' in df.columns:
                arabic_column = 'arabic_text'
            elif 'hadith_arabic' in df.columns:
                arabic_column = 'hadith_arabic'
            else:
                # Fallback to any column with 'arabic' in the name
                for col in df.columns:
                    if 'arabic' in str(col).lower():
                        arabic_column = col
                        break
        
        # If still no Arabic column found, return an error
        if arabic_column is None or arabic_column not in df.columns:
            return jsonify({
                'status': 'error', 
                'message': 'No Arabic text column found. Available columns: ' + ', '.join(df.columns)
            })
        
        # Convert row_idx to integer
        try:
            row_idx = int(row_idx)
            
            # Directly use the row index as provided (excel row number - 2)
            df_row_idx = row_idx - 2
            
            print(f"Requested row: {row_idx}, DataFrame index: {df_row_idx}, Max rows: {len(df)}")
            
            if df_row_idx < 0 or df_row_idx >= len(df):
                return jsonify({
                    'status': 'error', 
                    'message': f'Row index {row_idx} out of range (should be between 2 and {len(df)+1})'
                })
            
            # Get the Arabic text from the row
            arabic_text = df.iloc[df_row_idx][arabic_column]
            
            # Handle NaN or None values
            if pd.isna(arabic_text):
                arabic_text = "لا يوجد نص عربي" # "No Arabic text available" in Arabic
            else:
                arabic_text = str(arabic_text)
            
            return jsonify({
                'status': 'success',
                'arabic_text': arabic_text,
                'row_used': row_idx
            })
            
        except ValueError:
            return jsonify({'status': 'error', 'message': f'Invalid row index: {row_idx}'})
        
    except Exception as e:
        import traceback
        print(f"Error retrieving Arabic text: {str(e)}")
        print(traceback.format_exc())
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/translate_arabic_to_bangla', methods=['GET'])
def translate_arabic_to_bangla():
    try:
        row_idx = request.args.get('row_idx', type=int)
        
        if row_idx is None:
            return jsonify({'status': 'error', 'message': 'Row index is required'})

        input_file = get_input_file_path()
        if not os.path.exists(input_file):
            return jsonify({'status': 'error', 'message': 'Input file not found'})

        # Read the Excel file
        xls = pd.ExcelFile(input_file, engine='openpyxl')
        sheet_name = get_sheet_name()
        sheet_name = sheet_name if sheet_name in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_name)

        # Get the Arabic column name from config
        arabic_column = get_column_name('arabic_text')
        
        # Check for Arabic column
        if arabic_column not in df.columns:
            if 'arabic_text' in df.columns:
                arabic_column = 'arabic_text'
            elif 'hadith_arabic' in df.columns:
                arabic_column = 'hadith_arabic'
            else:
                for col in df.columns:
                    if 'arabic' in str(col).lower():
                        arabic_column = col
                        break
        
        if arabic_column is None or arabic_column not in df.columns:
            return jsonify({
                'status': 'error',
                'message': 'No Arabic text column found. Available columns: ' + ', '.join(df.columns)
            })

        # Validate row index
        df_row_idx = row_idx - 2  # Adjust for 0-based indexing and header row
        if df_row_idx < 0 or df_row_idx >= len(df):
            return jsonify({
                'status': 'error',
                'message': f'Row index {row_idx} out of range (should be between 2 and {len(df)+1})'
            })

        # Get the Arabic text
        arabic_text = df.iloc[df_row_idx][arabic_column]
        arabic_text = str(arabic_text) if not pd.isna(arabic_text) else "لا يوجد نص عربي"

        # Prepare the translation query
        from src.prompt import inject_variables, translate_arabic_to_bangla_prompt
        query = inject_variables(translate_arabic_to_bangla_prompt, {
            "arabic_text": arabic_text
        })

        # Call the AI model for translation
       
        translated_text = ask(query).text.strip()

        return jsonify({
            'status': 'success',
            'arabic_text': arabic_text,
            'translated_bangla': translated_text,
            'row_used': row_idx
        })

    except Exception as e:
        import traceback
        print(f"Error translating Arabic to Bangla: {str(e)}")
        print(traceback.format_exc())
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/recalculate_ratios', methods=['POST'])
def recalculate_ratios():
    input_file = get_input_file_path()
    if not os.path.exists(input_file):
        return jsonify({'status': 'error', 'message': 'Excel file not found'})
        
    try:
        # Get sheet and column names from config
        sheet_name = get_sheet_name()
        primary_text_col = get_column_name('primary_text')
        secondary_text_col = get_column_name('secondary_text')
        ratio_col = get_column_name('ratio')
        
        # Read the Excel file
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        
        # Calculate ratios for each row
        df[ratio_col] = df.apply(lambda row: difflib.SequenceMatcher(
            None, 
            extract_standard_letters(str(row[primary_text_col]) if pd.notna(row[primary_text_col]) else ""),
            extract_standard_letters(str(row[secondary_text_col]) if pd.notna(row[secondary_text_col]) else ""),
            autojunk=False
        ).ratio() * 100, axis=1)
        
        # Save the updated ratios back to Excel
        wb = load_workbook(input_file)
        if sheet_name not in wb.sheetnames:
            return jsonify({'status': 'error', 'message': f"'{sheet_name}' sheet not found in Excel file"})
            
        ws = wb[sheet_name]
        
        # Find the last column index or the existing ratio column
        ratio_col_idx = None
        last_col_idx = len(next(ws.rows))
        
        for idx, cell in enumerate(next(ws.rows)):
            if cell.value == ratio_col:
                ratio_col_idx = idx
                break
                
        ratio_col_letter = get_column_letter(ratio_col_idx + 1) if ratio_col_idx is not None else get_column_letter(last_col_idx + 1)
        
        # Add or update ratio header if needed
        if ratio_col_idx is None:
            ws[f'{ratio_col_letter}1'] = ratio_col
        
        # Update ratio values for each row
        for idx, ratio in enumerate(df[ratio_col], start=2):
            ws[f'{ratio_col_letter}{idx}'] = ratio
        
        wb.save(input_file)
        
        return jsonify({'status': 'success', 'message': 'Ratios recalculated successfully'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/select_chunk', methods=['POST'])
def select_chunk():
    global current_chunk
    chunk_path = request.form.get('chunk_path')
    
    # Validate the chunk path
    if chunk_path and os.path.exists(chunk_path) and os.path.isfile(chunk_path):
        current_chunk = chunk_path
        print(f"Selected chunk changed to: {current_chunk}")
    else:
        print(f"Warning: Invalid chunk path: {chunk_path}")
    
    # Redirect back to the index page with the same query parameters
    return redirect(url_for('index', **request.args))

@app.context_processor
def utility_processor():
    return dict(urlencode=urlencode)

@app.route('/regenerate_with_prompt_1', methods=['POST'])
def regenerate_with_prompt_1():
    row_idx = request.form.get('row_idx', type=int)
    try:
        input_file = get_input_file_path()
        
        # 1. Read row data
        from src.generate_cell import read_row
        arabic_text, _, current_bangla = read_row(row_idx, input_file)
        
        # 2. Generate text using prompt 1
        from src.prompt import inject_variables, read_file
        from src.ai import ask
        query = inject_variables(read_file("./prompts/1.md"), {
            "hadis_arabic_text": arabic_text,
            "previous_generated_bangla": current_bangla
        })
        new_text = ask(query).text.strip()

        # 3. Update Excel file with new text
        if not os.path.exists(input_file):
            return jsonify({'status': 'error', 'message': 'Excel file not found after regeneration'})

        wb = load_workbook(input_file)
        sheet_name = get_sheet_name()
        if sheet_name not in wb.sheetnames:
            return jsonify({'status': 'error', 'message': f'{sheet_name} sheet not found in Excel file'})
        ws = wb[sheet_name]

        header = next(ws.rows)
        primary_text_col_name = get_column_name('primary_text')
        secondary_text_col_name = get_column_name('secondary_text')
        
        secondary_text_col_idx = 1
        primary_text_col_idx = 0
        
        for idx, cell in enumerate(header):
            if cell.value == secondary_text_col_name:
                secondary_text_col_idx = idx
            elif cell.value == primary_text_col_name:
                primary_text_col_idx = idx

        excel_row = row_idx + 2
        cell_address = f'{get_column_letter(secondary_text_col_idx + 1)}{excel_row}'
        ws[cell_address].value = new_text
        ws[cell_address].fill = PatternFill(fill_type=None)  # Clear existing fill

        wb.save(input_file)

        # 4. Get updated color status and prepare comparison data
        color_status = get_cell_color_status()
        row_approval = color_status.get(excel_row, {'col_b': False, 'col_b_type': None})
        col_b_approved = row_approval['col_b']
        col_b_type = row_approval['col_b_type']

        # Get original text from Column A for comparison
        col_a_cell = ws[f'{get_column_letter(primary_text_col_idx + 1)}{excel_row}']
        col_a_text = str(col_a_cell.value) if col_a_cell.value is not None else ''
        highlighted_a, highlighted_b, status = compare_text(col_a_text, new_text)

        return jsonify({
            'status': 'success',
            'new_text': new_text,
            'highlighted_html': highlighted_b,
            'highlighted_a_html': highlighted_a,  # Include highlighted HTML for column A
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


@app.route('/regenerate_with_prompt_2', methods=['POST'])
def regenerate_with_prompt_2():
    row_idx = request.form.get('row_idx', type=int)
    try:
        input_file = get_input_file_path()
        
        # 1. Read row data
        from src.generate_cell import read_row
        arabic_text, _, current_bangla = read_row(row_idx, input_file)
        
        # 2. Generate text using prompt 2
        from src.prompt import inject_variables, read_file
        from src.ai import ask
        query = inject_variables(read_file("./prompts/2.md"), {
            "hadis_arabic_text": arabic_text,
            "previous_generated_bangla": current_bangla
        })
        new_text = ask(query).text.strip()

        # 3. Update Excel file with new text
        if not os.path.exists(input_file):
            return jsonify({'status': 'error', 'message': 'Excel file not found after regeneration'})

        wb = load_workbook(input_file)
        sheet_name = get_sheet_name()
        if sheet_name not in wb.sheetnames:
            return jsonify({'status': 'error', 'message': f'{sheet_name} sheet not found in Excel file'})
        ws = wb[sheet_name]

        header = next(ws.rows)
        primary_text_col_name = get_column_name('primary_text')
        secondary_text_col_name = get_column_name('secondary_text')
        
        secondary_text_col_idx = 1
        primary_text_col_idx = 0
        
        for idx, cell in enumerate(header):
            if cell.value == secondary_text_col_name:
                secondary_text_col_idx = idx
            elif cell.value == primary_text_col_name:
                primary_text_col_idx = idx

        excel_row = row_idx + 2
        cell_address = f'{get_column_letter(secondary_text_col_idx + 1)}{excel_row}'
        ws[cell_address].value = new_text
        ws[cell_address].fill = PatternFill(fill_type=None)  # Clear existing fill

        wb.save(input_file)

        # 4. Get updated color status and prepare comparison data
        color_status = get_cell_color_status()
        row_approval = color_status.get(excel_row, {'col_b': False, 'col_b_type': None})
        col_b_approved = row_approval['col_b']
        col_b_type = row_approval['col_b_type']

        # Get original text from Column A for comparison
        col_a_cell = ws[f'{get_column_letter(primary_text_col_idx + 1)}{excel_row}']
        col_a_text = str(col_a_cell.value) if col_a_cell.value is not None else ''
        highlighted_a, highlighted_b, status = compare_text(col_a_text, new_text)

        return jsonify({
            'status': 'success',
            'new_text': new_text,
            'highlighted_html': highlighted_b,
            'highlighted_a_html': highlighted_a,  # Include highlighted HTML for column A
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


@app.route('/regenerate_multiple_with_prompt_1', methods=['POST'])
def regenerate_multiple_with_prompt_1():
    row_ids = request.json.get('row_ids', [])

    print(f"Regenerating rows with prompt 1: {row_ids}")
    
    if not row_ids:
        return jsonify({'status': 'error', 'message': 'No row IDs provided'})
    
    try:
        import concurrent.futures
        
        input_file = get_input_file_path()
        results = []
        
        if not os.path.exists(input_file):
            return jsonify({'status': 'error', 'message': 'Excel file not found'})
        
        # First, generate all the new texts in parallel
        generated_texts = {}
        
        def generate_text_for_row_prompt_1(row_idx):
            try:
                print(f"Generating text with prompt 1 for row: {row_idx}")
                
                # Read row data
                from src.generate_cell import read_row
                arabic_text, _, current_bangla = read_row(row_idx, input_file)
                
                # Generate text using prompt 1
                from src.prompt import inject_variables, read_file
                from src.ai import ask
                query = inject_variables(read_file("./prompts/1.md"), {
                    "hadis_arabic_text": arabic_text,
                    "previous_generated_bangla": current_bangla
                })
                new_text = ask(query).text.strip()
                
                return {'status': 'success', 'row_idx': row_idx, 'new_text': new_text}
            except Exception as e:
                import traceback
                return {
                    'status': 'error',
                    'row_idx': row_idx,
                    'message': str(e),
                    'traceback': traceback.format_exc()
                }
        
        # Generate all texts in parallel
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(10, len(row_ids))) as executor:
            # Submit all generation tasks
            future_to_row = {executor.submit(generate_text_for_row_prompt_1, row_idx): row_idx for row_idx in row_ids}
            
            # Collect results as they complete
            for future in concurrent.futures.as_completed(future_to_row):
                row_idx = future_to_row[future]
                try:
                    result = future.result()
                    if result['status'] == 'success':
                        generated_texts[row_idx] = result['new_text']
                    else:
                        results.append(result)  # Store error results
                except Exception as e:
                    results.append({
                        'status': 'error',
                        'row_idx': row_idx,
                        'message': str(e)
                    })
        
        # If there are successful generations, update the Excel file only once
        if generated_texts:
            try:
                # Load workbook once
                wb = load_workbook(input_file)
                sheet_name = get_sheet_name()
                
                if sheet_name not in wb.sheetnames:
                    return jsonify({'status': 'error', 'message': f'{sheet_name} sheet not found'})
                
                ws = wb[sheet_name]
                
                # Find column indices
                header = next(ws.rows)
                primary_text_col_name = get_column_name('primary_text')
                secondary_text_col_name = get_column_name('secondary_text')
                
                secondary_text_col_idx = 1
                primary_text_col_idx = 0
                
                for idx, cell in enumerate(header):
                    if cell.value == secondary_text_col_name:
                        secondary_text_col_idx = idx
                    elif cell.value == primary_text_col_name:
                        primary_text_col_idx = idx
                
                # Update all cells at once
                for row_idx, new_text in generated_texts.items():
                    excel_row = row_idx + 2
                    cell_address = f'{get_column_letter(secondary_text_col_idx + 1)}{excel_row}'
                    ws[cell_address].value = new_text
                    ws[cell_address].fill = PatternFill(fill_type=None)  # Clear existing fill
                
                # Save the workbook once after all updates
                wb.save(input_file)
                
                # Now collect all comparison results
                for row_idx, new_text in generated_texts.items():
                    excel_row = row_idx + 2
                    
                    # Get original text for comparison
                    col_a_cell = ws[f'{get_column_letter(primary_text_col_idx + 1)}{excel_row}']
                    col_a_text = str(col_a_cell.value) if col_a_cell.value is not None else ''
                    highlighted_a, highlighted_b, status = compare_text(col_a_text, new_text)
                    
                    # Fetch color status for this row
                    color_status = get_cell_color_status()
                    row_approval = color_status.get(excel_row, {'col_b': False, 'col_b_type': None})
                    col_b_approved = row_approval['col_b']
                    col_b_type = row_approval['col_b_type']
                    
                    results.append({
                        'status': 'success',
                        'row_idx': row_idx,
                        'new_text': new_text,
                        'highlighted_html': highlighted_b,
                        'highlighted_a_html': highlighted_a,
                        'diff_status': status,
                        'col_b_approved': col_b_approved,
                        'col_b_type': col_b_type
                    })
                
            except Exception as e:
                import traceback
                error_message = f"Error updating Excel file: {str(e)}"
                traceback.print_exc()
                
                # Add error for each row that was not already recorded as an error
                for row_idx in generated_texts.keys():
                    if not any(r.get('row_idx') == row_idx and r.get('status') == 'error' for r in results):
                        results.append({
                            'status': 'error',
                            'row_idx': row_idx,
                            'message': error_message
                        })
        
        # Return all results
        return jsonify({
            'status': 'success',
            'message': f'Completed regeneration with prompt 1 for {len(generated_texts)} rows. {len(row_ids) - len(generated_texts)} failed.',
            'results': results
        })
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({
            'status': 'error',
            'message': f'An unexpected error occurred: {str(e)}'
        })


@app.route('/regenerate_multiple_with_prompt_2', methods=['POST'])
def regenerate_multiple_with_prompt_2():
    row_ids = request.json.get('row_ids', [])

    print(f"Regenerating rows with prompt 2: {row_ids}")
    
    if not row_ids:
        return jsonify({'status': 'error', 'message': 'No row IDs provided'})
    
    try:
        import concurrent.futures
        
        input_file = get_input_file_path()
        results = []
        
        if not os.path.exists(input_file):
            return jsonify({'status': 'error', 'message': 'Excel file not found'})
        
        # First, generate all the new texts in parallel
        generated_texts = {}
        
        def generate_text_for_row_prompt_2(row_idx):
            try:
                print(f"Generating text with prompt 2 for row: {row_idx}")
                
                # Read row data
                from src.generate_cell import read_row
                arabic_text, _, current_bangla = read_row(row_idx, input_file)
                
                # Generate text using prompt 2
                from src.prompt import inject_variables, read_file
                from src.ai import ask
                query = inject_variables(read_file("./prompts/2.md"), {
                    "hadis_arabic_text": arabic_text,
                    "previous_generated_bangla": current_bangla
                })
                new_text = ask(query).text.strip()
                
                return {'status': 'success', 'row_idx': row_idx, 'new_text': new_text}
            except Exception as e:
                import traceback
                return {
                    'status': 'error',
                    'row_idx': row_idx,
                    'message': str(e),
                    'traceback': traceback.format_exc()
                }
        
        # Generate all texts in parallel
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(10, len(row_ids))) as executor:
            # Submit all generation tasks
            future_to_row = {executor.submit(generate_text_for_row_prompt_2, row_idx): row_idx for row_idx in row_ids}
            
            # Collect results as they complete
            for future in concurrent.futures.as_completed(future_to_row):
                row_idx = future_to_row[future]
                try:
                    result = future.result()
                    if result['status'] == 'success':
                        generated_texts[row_idx] = result['new_text']
                    else:
                        results.append(result)  # Store error results
                except Exception as e:
                    results.append({
                        'status': 'error',
                        'row_idx': row_idx,
                        'message': str(e)
                    })
        
        # If there are successful generations, update the Excel file only once
        if generated_texts:
            try:
                # Load workbook once
                wb = load_workbook(input_file)
                sheet_name = get_sheet_name()
                
                if sheet_name not in wb.sheetnames:
                    return jsonify({'status': 'error', 'message': f'{sheet_name} sheet not found'})
                
                ws = wb[sheet_name]
                
                # Find column indices
                header = next(ws.rows)
                primary_text_col_name = get_column_name('primary_text')
                secondary_text_col_name = get_column_name('secondary_text')
                
                secondary_text_col_idx = 1
                primary_text_col_idx = 0
                
                for idx, cell in enumerate(header):
                    if cell.value == secondary_text_col_name:
                        secondary_text_col_idx = idx
                    elif cell.value == primary_text_col_name:
                        primary_text_col_idx = idx
                
                # Update all cells at once
                for row_idx, new_text in generated_texts.items():
                    excel_row = row_idx + 2
                    cell_address = f'{get_column_letter(secondary_text_col_idx + 1)}{excel_row}'
                    ws[cell_address].value = new_text
                    ws[cell_address].fill = PatternFill(fill_type=None)  # Clear existing fill
                
                # Save the workbook once after all updates
                wb.save(input_file)
                
                # Now collect all comparison results
                for row_idx, new_text in generated_texts.items():
                    excel_row = row_idx + 2
                    
                    # Get original text for comparison
                    col_a_cell = ws[f'{get_column_letter(primary_text_col_idx + 1)}{excel_row}']
                    col_a_text = str(col_a_cell.value) if col_a_cell.value is not None else ''
                    highlighted_a, highlighted_b, status = compare_text(col_a_text, new_text)
                    
                    # Fetch color status for this row
                    color_status = get_cell_color_status()
                    row_approval = color_status.get(excel_row, {'col_b': False, 'col_b_type': None})
                    col_b_approved = row_approval['col_b']
                    col_b_type = row_approval['col_b_type']
                    
                    results.append({
                        'status': 'success',
                        'row_idx': row_idx,
                        'new_text': new_text,
                        'highlighted_html': highlighted_b,
                        'highlighted_a_html': highlighted_a,
                        'diff_status': status,
                        'col_b_approved': col_b_approved,
                        'col_b_type': col_b_type
                    })
                
            except Exception as e:
                import traceback
                error_message = f"Error updating Excel file: {str(e)}"
                traceback.print_exc()
                
                # Add error for each row that was not already recorded as an error
                for row_idx in generated_texts.keys():
                    if not any(r.get('row_idx') == row_idx and r.get('status') == 'error' for r in results):
                        results.append({
                            'status': 'error',
                            'row_idx': row_idx,
                            'message': error_message
                        })
        
        # Return all results
        return jsonify({
            'status': 'success',
            'message': f'Completed regeneration with prompt 2 for {len(generated_texts)} rows. {len(row_ids) - len(generated_texts)} failed.',
            'results': results
        })
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({
            'status': 'error',
            'message': f'An unexpected error occurred: {str(e)}'
        })

@app.route('/regenerate_with_custom_prompt', methods=['POST'])
def regenerate_with_custom_prompt():
    row_idx = request.form.get('row_idx', type=int)
    custom_prompt = request.form.get('prompt', '')
    
    if not custom_prompt.strip():
        return jsonify({'status': 'error', 'message': 'Empty prompt provided'})
    
    try:
        input_file = get_input_file_path()
        
        # 1. Read row data
        from src.generate_cell import read_row
        arabic_text, col_a_text, col_b_text = read_row(row_idx, input_file)
        
        # 2. Process the custom prompt by replacing placeholders
        processed_prompt = inject_variables(custom_prompt, {
            "arabic_text": arabic_text,
            "col_a_text": col_a_text,
            "col_b_text": col_b_text
        })
        
        # 3. Generate text using the custom prompt
        from src.ai import ask
        new_text = ask(processed_prompt).text.strip()

        # 4. Update Excel file with new text
        if not os.path.exists(input_file):
            return jsonify({'status': 'error', 'message': 'Excel file not found after regeneration'})

        wb = load_workbook(input_file)
        sheet_name = get_sheet_name()
        if sheet_name not in wb.sheetnames:
            return jsonify({'status': 'error', 'message': f'{sheet_name} sheet not found in Excel file'})
        ws = wb[sheet_name]

        header = next(ws.rows)
        primary_text_col_name = get_column_name('primary_text')
        secondary_text_col_name = get_column_name('secondary_text')
        
        secondary_text_col_idx = 1
        primary_text_col_idx = 0
        
        for idx, cell in enumerate(header):
            if cell.value == secondary_text_col_name:
                secondary_text_col_idx = idx
            elif cell.value == primary_text_col_name:
                primary_text_col_idx = idx

        excel_row = row_idx + 2
        cell_address = f'{get_column_letter(secondary_text_col_idx + 1)}{excel_row}'
        ws[cell_address].value = new_text
        ws[cell_address].fill = PatternFill(fill_type=None)  # Clear existing fill

        wb.save(input_file)

        # 5. Get updated color status and prepare comparison data
        color_status = get_cell_color_status()
        row_approval = color_status.get(excel_row, {'col_b': False, 'col_b_type': None})
        col_b_approved = row_approval['col_b']
        col_b_type = row_approval['col_b_type']

        # Get original text from Column A for comparison
        col_a_cell = ws[f'{get_column_letter(primary_text_col_idx + 1)}{excel_row}']
        col_a_text = str(col_a_cell.value) if col_a_cell.value is not None else ''
        highlighted_a, highlighted_b, status = compare_text(col_a_text, new_text)

        return jsonify({
            'status': 'success',
            'new_text': new_text,
            'highlighted_html': highlighted_b,
            'highlighted_a_html': highlighted_a,
            'diff_status': status,
            'col_b_approved': col_b_approved,
            'col_b_type': col_b_type
        })

    except FileNotFoundError as e:
        return jsonify({'status': 'error', 'message': str(e)})
    except ValueError as e:
        return jsonify({'status': 'error', 'message': str(e)})
    except Exception as e:
        print(f"Error during custom prompt regeneration for row {row_idx}: {type(e).__name__} - {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'status': 'error', 'message': f'An unexpected error occurred: {str(e)}'})

if __name__ == '__main__': app.run(debug=True,port=8000)
