#!/usr/bin/env python3
import os
import re
import sys
import argparse
from openpyxl import load_workbook

def split_excel(input_file, output_dir='chunks', rows_per_chunk=500):
    """
    Split a large Excel file into smaller chunks without style preservation.
    
    Args:
        input_file (str): Path to the input Excel file
        output_dir (str): Directory to save the chunks
        rows_per_chunk (int): Maximum number of rows per chunk
    
    Returns:
        list: List of generated chunk files
    """
    # Ensure output directory exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Load the workbook
    print(f"Loading workbook: {input_file}")
    wb = load_workbook(input_file)
    
    # Get the active sheet (or first sheet)
    sheet_name = 'hadith' if 'hadith' in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]
    
    # Get total rows (excluding header)
    total_rows = ws.max_row - 1  # Subtract 1 for header row
    
    # Calculate number of chunks needed
    num_chunks = (total_rows + rows_per_chunk - 1) // rows_per_chunk
    
    print(f"Total rows: {total_rows}")
    print(f"Rows per chunk: {rows_per_chunk}")
    print(f"Number of chunks: {num_chunks}")
    
    chunk_files = []
    
    # Process each chunk
    for chunk_idx in range(num_chunks):
        # Calculate row range for this chunk
        start_row = chunk_idx * rows_per_chunk + 1  # +1 because row 1 is the header
        end_row = min((chunk_idx + 1) * rows_per_chunk, total_rows) + 1  # +1 because row 1 is the header
        
        # Create a new workbook for the chunk
        chunk_wb = load_workbook(filename=input_file)
        chunk_ws = chunk_wb[sheet_name]
        
        # Remove rows outside of our chunk
        # First, remove rows after our chunk (higher row numbers)
        for row_idx in range(chunk_ws.max_row, end_row, -1):
            chunk_ws.delete_rows(row_idx)
        
        # Then, remove rows before our chunk (lower row numbers, but keep header)
        for row_idx in range(start_row-1, 1, -1):
            chunk_ws.delete_rows(row_idx)
        
        # Define chunk filename
        chunk_filename = f"chunk_{chunk_idx+1}_rows_{start_row}-{end_row-1}.xlsx"
        chunk_path = os.path.join(output_dir, chunk_filename)
        
        # Save the chunk
        print(f"Saving chunk {chunk_idx+1}/{num_chunks}: {chunk_path}")
        chunk_wb.save(chunk_path)
        chunk_files.append(chunk_path)
    
    print(f"Splitting complete. Created {len(chunk_files)} chunks.")
    return chunk_files

def merge_excel(chunk_dir='chunks', output_file=None):
    """
    Merge chunked Excel files back into a single file without style preservation.
    
    Args:
        chunk_dir (str): Directory containing the chunk files
        output_file (str): Output file path. If None, will be 'merged_output.xlsx'
    
    Returns:
        str: Path to the merged file
    """
    # Find all chunk files
    chunk_files = []
    chunk_pattern = re.compile(r'chunk_(\d+)_rows_(\d+)-(\d+)\.xlsx')
    
    for filename in os.listdir(chunk_dir):
        if chunk_pattern.match(filename):
            chunk_match = chunk_pattern.match(filename)
            chunk_num = int(chunk_match.group(1))
            
            chunk_files.append({
                'filename': os.path.join(chunk_dir, filename),
                'chunk_num': chunk_num
            })
    
    # Sort by chunk number
    chunk_files.sort(key=lambda x: x['chunk_num'])
    
    if not chunk_files:
        print("No chunk files found")
        return None
    
    # Load the first chunk as our base workbook
    print(f"Loading first chunk: {chunk_files[0]['filename']}")
    merged_wb = load_workbook(chunk_files[0]['filename'])
    
    sheet_name = 'hadith' if 'hadith' in merged_wb.sheetnames else merged_wb.sheetnames[0]
    merged_ws = merged_wb[sheet_name]
    
    # Current row count in the merged workbook
    current_row_count = merged_ws.max_row
    
    # Process remaining chunks
    for i, chunk_info in enumerate(chunk_files[1:], 1):
        print(f"Processing chunk {i+1}/{len(chunk_files)}: {chunk_info['filename']}")
        
        # Load the chunk
        chunk_wb = load_workbook(chunk_info['filename'])
        chunk_ws = chunk_wb[sheet_name]
        
        # Copy rows from chunk to merged workbook (skip header row)
        for row_idx in range(2, chunk_ws.max_row + 1):
            # Copy row to merged workbook
            merged_ws.insert_rows(current_row_count + 1)
            current_row_count += 1
            
            # Copy cell values only
            for col_idx in range(1, chunk_ws.max_column + 1):
                source_cell = chunk_ws.cell(row=row_idx, column=col_idx)
                target_cell = merged_ws.cell(row=current_row_count, column=col_idx)
                target_cell.value = source_cell.value
    
    # Save the merged workbook
    if output_file is None:
        output_file = 'merged_output.xlsx'
    
    print(f"Saving merged file: {output_file}")
    merged_wb.save(output_file)
    
    print(f"Merge complete. Total rows: {current_row_count}")
    return output_file

def main():
    parser = argparse.ArgumentParser(description='Split and merge Excel files')
    parser.add_argument('--action', choices=['split', 'merge'], required=True, help='Action to perform (split or merge)')
    parser.add_argument('--input', help='Input Excel file to split')
    parser.add_argument('--output', help='Output file or directory')
    parser.add_argument('--chunk-dir', default='chunks', help='Directory for chunks (default: chunks)')
    parser.add_argument('--rows', type=int, default=500, help='Rows per chunk (default: 500)')
    
    args = parser.parse_args()
    
    if args.action == 'split':
        if not args.input:
            print("Error: --input is required for split action")
            sys.exit(1)
            
        output_dir = args.output or args.chunk_dir
        split_excel(args.input, output_dir, args.rows)
        
    elif args.action == 'merge':
        chunk_dir = args.chunk_dir
        output_file = args.output
        merge_excel(chunk_dir, output_file)

if __name__ == '__main__':
    main()
