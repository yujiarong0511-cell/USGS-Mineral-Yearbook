import pandas as pd
import os
import sys
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.append(parent_dir)
from algorithm import get_consecutive_ranges
from getcolwidth import get_real_column_width
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter 
import xlwings as xw
import re

def structuring_tables(continent_select, country_year):
    # %%
    # Create a list of file names referring to only tables 2019 MYB. 
    # continent_select = "southamerica" # [INPUT VAR 1]
    input_dir = "../../data/original_table/usgs_" + continent_select + "/"
    # input_filenames_all = os.listdir(input_dir)
    # input_filenames_2019 = [name for name in input_filenames_all if "2019" in name]
    output_dir = "../../data/structured_tables/temp/usgs_" + continent_select + "/" # Temp files

    # %% [markdown]
    # ##### Take Single Country (2019) as an example (encapsulated into .py)

    # %%
    name = country_year + ".xlsx" # [INPUT VAR 2]

    # %% [markdown]
    # ##### 1.1 Identify redundant rows to remove
    # *Important note: All rows must be remove at the same time, otherwise column index should be reloaded.*

    # %%
    # Read-in Table
    input_filepath =  os.path.join(input_dir, name)
    print(f"Input: {input_filepath}")
    fullfile_openpyxl = load_workbook(input_filepath)
    # Check if `fullfile_openpyxl.sheetnames` includes "Table 2" or "Table2"
    table_openpyxl = list(fullfile_openpyxl)[2] # DEBUG: Table 2 are at times not named as 'Table 2' in Excel (e.g., in China-2019 as 'Table2')

    # %% [markdown]
    # (I) Penetrating merged cells (NOTE: This block is provisonally abandoned, substituted by the one below)

    # %%
    # DEBUG: We remove all rows where there exist merged cells (rather than solely cells penetrating all columns)
    '''
    merged_cells = list(table_openpyxl.merged_cells.ranges)
    merged_cells = sorted(merged_cells, key=lambda r: (r.min_row, r.min_col)) # DEBUG: Return of `table_openpyxl.merged_cells.ranges` could at times out of order (e.g., Urugary-2019)

    ## Delete rows where merged cells penetrated through all columns
    table_maxcol = merged_cells[0].bounds[2] # Determinator of penetrating merged cells
    merged_cells_penetrate = [mc for mc in merged_cells if mc.bounds[2] == table_maxcol] # Select all penetrating merged cells
    merged_cells_penetrate_rows = [mc_row.bounds[1] for mc_row in merged_cells_penetrate] # Locate rows of penetrating merged cells
    merged_cells_penetrate_rows.sort(reverse=False)

    print(merged_cells_penetrate_rows)
    print(f"table_maxcol: {table_maxcol}")
    '''

    # %% [markdown]
    # (I*) All merged cells (NOTE: This block replace the previous one)

    # %%
    # DEBUG: We remove all rows where there exist merged cells (rather than solely cells penetrating all columns)
    merged_cells = list(table_openpyxl.merged_cells.ranges)
    merged_cells = sorted(merged_cells, key=lambda r: (r.min_row, r.min_col)) # DEBUG: Return of `table_openpyxl.merged_cells.ranges` could at times out of order (e.g., Urugary-2019)
    table_maxcol = merged_cells[0].bounds[2] # Determinator of penetrating merged cells (DEBUG: While it's not used in this block, it serves as input for subsequent ones)

    merged_cells_rows = [mc_row.bounds[1] for mc_row in merged_cells]
    merged_cells_rows.sort(reverse=False)
    print(merged_cells_rows)

    # %% [markdown]
    # (II) Footnotes (Criteria: below the last lower border)

    # %%
    last_row_with_lowerborder = None
    # IMPORTANT DEBUG: In case the column range is too broad (extending to 1000+)
    col_range = table_openpyxl.max_column if table_openpyxl.max_column < 10 else 10
    for row in range(table_openpyxl.max_row, 0, -1): # Iterate bottom-up
        for col in range(1, col_range+1):
            cell = table_openpyxl.cell(row=row, column=col) # DEBUG: Polish the way we invoke cells
            next_cell = table_openpyxl.cell(row=row+1, column=col)
            if cell.border.bottom.style or next_cell.border.top.style: # IMPORTANT DEBUG: Check border style from both side
                last_row_with_lowerborder = row + 1 # IMPORTANT: First footnote row
            '''
            if cell.border.bottom.style: # Stop the loop as soon as the first cell with lower border is found.

                footnote_identify_regex = r"^(eEstimate|1[a-zA-Z])"
                next_cell_value = table_openpyxl[row+1][0].value
                if re.match(footnote_identify_regex, next_cell_value): # Determine if the next row is footnote
                    last_row_with_lowerborder = row + 1
                else: # VERY IMPORTANT: cell.border.bottom.style can sometimes fail to identify frame layout of cells whose subsequent row contains merged cells.
                    last_row_with_lowerborder = row + 2 # [PROVISIONAL]
            '''
        if last_row_with_lowerborder is not None:
            break
    if last_row_with_lowerborder is not None:
        footnote_rows = list(range(last_row_with_lowerborder,table_openpyxl.max_row+1))
    else:
        footnote_rows = []
        print("No footnote identified in this table! (Make sure to recheck the layout)")

    print(row, last_row_with_lowerborder, table_openpyxl.max_row)
    print(footnote_rows)

    # %% [markdown]
    # (III) Repeated headers

    # %%
    headers_rows = []
    for row in table_openpyxl.iter_rows(min_col=1, max_col=1):
        cell = row[0]
        row_num = row[0].row
        if cell.value and "commodit" in str(cell.value).lower(): # DEBUG: Sometimes "Country and commodity" could appear in header (e.g., Urugary-2019)
            # IMPORTANT DEBUG: Two-sided border style check (with first row naturally has its top border)
            if cell.border.top.style is not None or (row_num > 1 and table_openpyxl[row_num-1][0].border.bottom.style is not None) or row_num == 1:
                headers_rows.append(row_num)
            else:
                row_num_start = 1
                for r in range(row_num-1, 0, -1): # DEBUG: Replace row_num with row_num-1 to 
                    previous_cell = table_openpyxl[r][0]
                    if previous_cell.border.top.style:
                        row_num_start = r
                        break
                    # IMPORTANT DEBUG: Borders features are often saved only in one of the two adjacent cells, meaning we could at times miss the cell if there's only one-sided check.
                    if previous_cell.border.bottom.style:
                        row_num_start = r+1
                        break
                headers_rows = headers_rows + list(range(row_num_start, row_num+1))
    print(headers_rows)

    # %% [markdown]
    # (IV) Containing "—" or "Continued"

    # %%
    # IMPORTANT DEBUG: "—" and "Continued" should be removed WITHIN cells but the rows themselves should be kept. Otherwise certain entries would be mistakenly removed.
    '''
    continue_rows = []
    for row in table_openpyxl.iter_rows(min_col=1, max_col=1):
        cell = row[0]
        row_num = row[0].row
        if cell.value and ("—" in str(cell.value) or "continue" in str(cell.value).lower()):
            continue_rows.append(row_num)

    print(continue_rows)
    '''

    # %% [markdown]
    # (V) All upper letter (excluding "—" and "Continue(d)") in `Commodity` column

    # %%
    import re
    allupper_rows = []
    clean_regex = re.compile(r'—|continued?', re.IGNORECASE) # Construct regular expression to match keywords
    for row in table_openpyxl.iter_rows(min_col=1, max_col=1):
        cell = row[0]
        row_num = row[0].row
        cell_value = row[0].value # Get value from the cell (return None the cell is empty)
        if cell_value and isinstance(cell_value, str): 
            cleaned_cell_value = clean_regex.sub('', cell_value).strip() # Remove all "—" and "continue(d)" if they ever exist
            if cleaned_cell_value and cleaned_cell_value.isupper():
                allupper_rows.append(row_num)

    print(allupper_rows)

    # %% [markdown]
    # **Gather all redundant rows and remove repeated ones:**

    # %%
    # Gather all redundant rows and remove repeated ones (with `list(set(list_A + list_B))`)
    # DEBUG: (1)"continue_rows" is removed; (2) "merged_cells_rows" substitued "merged_cells_penetrate_rows"
    redundant_rows = list(set(merged_cells_rows + footnote_rows + headers_rows + allupper_rows))
    redundant_rows.sort(reverse=True) # Sort the rows in reverse order

    # Concatenate list of located rows into a single string (consecutive series are simplified with colon like "1:5")
    redundant_rows_concise = get_consecutive_ranges(redundant_rows) # Invoke function "get_consecutive_ranges(list)" from ../../module/structuring_tables/algorithm.py
    redundant_rows_concise_str = ",".join(string for string in redundant_rows_concise)

    print(redundant_rows_concise_str)

    # %% [markdown]
    # ##### Step 1.2: Identify isolation columns to remove

    # %% [markdown]
    # (I) Apparent isolation

    # %%
    # Delete columns significantly narrower than its right neighbor (isolation strips)
    # IMPORTANT: table_openpyxl.column_dimensions['D'].width can return false value 13.0 (default value at the back end of openpyxl when apparent width is not attainable, given width of adjacent columns are sometimes compressed in excel). Thus we encapsulated a module to decompress and reveal the real width.
    real_col_width_list = get_real_column_width(table_openpyxl)
    threshold_ratio = 0.3 ## NOTE: The threshold needs to be fine-tuned to tailor to layouts in different tables.
    isolation_col = []
    for col_idx in range(2, table_maxcol+1): # IMPORTANT: Start from the second column and end at the last column
        # Get reference letter (e.g., A/B/C...) of current column and its neighbors
        left_col_letter = get_column_letter(col_idx - 1)
        current_col_letter = get_column_letter(col_idx)
        right_col_letter = get_column_letter(col_idx + 1)

        # Get width of current column and its neighbors
        left_col_width = real_col_width_list[left_col_letter]
        current_col_width = real_col_width_list[current_col_letter]
        right_col_width = real_col_width_list[right_col_letter]

        if current_col_width < (right_col_width * threshold_ratio): # Compare solely with the right-hand column
            isolation_col.append(current_col_letter)

    print(isolation_col)

    # %% [markdown]
    # (II) Isolation strip between `Commodity` column and `Unit` column

    # %%
    # Identify totally empty columns
    empty_columns = []
    for col in list(table_openpyxl.columns)[:10]: # DEBUG: Sometimes "table_openpyxl.columns" can contain huge numbers of empty columns among which the majority are far from the right boundary of the table with contents
        is_empty = all(cell.value is None or str(cell.value).strip()=="" 
                    for cell in col)
        if is_empty:
            col_letter = get_column_letter(col[0].column)
            empty_columns.append(col_letter)
    print(empty_columns)

    # %% [markdown]
    # **Gather all redundant columns and remove repeated ones:**

    # %%
    redundant_cols = list(set(isolation_col + empty_columns))
    redundant_cols.sort(reverse=True) # IMPORTANT: Columns must be removed from right to left

    # Compile list of isolation strips into a single valid string comprehensible to xlwings package
    redundant_col_str = ','.join([f"{col}:{col}" for col in redundant_cols])

    print(redundant_col_str)

    # %% [markdown]
    # ##### Step 1.3: Operate removal

    # %%
    # To save time, we initiate a new Excel process only if Excel is not running at the back end
    if len(xw.apps) > 0:
        app = xw.apps.active
        print("Running Excel process is invoked :)")
    else:
        app = xw.App(visible=True)
        print("New Excel process has started!")

    # Open the original table
    fullfile_xlwings = app.books.open(input_filepath) # Open .xlsx file
    # table_xlwings = fullfile_xlwings.sheets['Table 2'] # Refer to worksheet 'Table 2'
    table_xlwings = fullfile_xlwings.sheets[2] # DEBUG: Table 2 are at times not named as 'Table 2' in Excel (e.g., in China-2019 as 'Table2')

    ## Set horizontal border style before removing rows (to avoid missing crucial boundaries)
    all_rows_object = table_xlwings.api.Range(redundant_rows_concise_str)

    # IMPORTANT DEBUG: Iterate through all non-contiguous row blocks (Areas) and reinforce boundaries
    for area in all_rows_object.Areas:
        # Get the first row number and total row count of the current continuous block
        first_row = area.Row
        row_count = area.Rows.Count
        last_row = first_row + row_count - 1
        
        # A. Reinforce the bottom border of the "retained row above"
        if first_row > 1:  # Ensure it's not the first row to avoid out-of-bounds error
            # Precisely target the exactly 1 row directly above
            above_row = area.Worksheet.Rows(first_row - 1)
            above_row.Borders(xw.constants.BordersIndex.xlEdgeBottom).LineStyle = xw.constants.LineStyle.xlContinuous
            above_row.Borders(xw.constants.BordersIndex.xlEdgeBottom).Weight = xw.constants.BorderWeight.xlThin

        # B. Reinforce the top border of the "retained row below"
        if last_row < 1048576:  # Ensure it doesn't exceed Excel's maximum row limit
            # Precisely target the exactly 1 row directly below
            below_row = area.Worksheet.Rows(last_row + 1)
            below_row.Borders(xw.constants.BordersIndex.xlEdgeTop).LineStyle = xw.constants.LineStyle.xlContinuous
            below_row.Borders(xw.constants.BordersIndex.xlEdgeTop).Weight = xw.constants.BorderWeight.xlThin

    ## Delete all rows (in batch) where merged cells penetrate through all columns
    all_rows_object.Delete() # Highly optimized (and no contraints on order of rows)

    ## Delete all columns (in batch) serving as isolation strips
    table_xlwings.api.Range(redundant_col_str).Delete() # [NEWLY ADDED]

    # Save the modified copy and close the original doc
    temp_name = name.split('.')[0] + "_1.xlsx"
    output_filepath = os.path.join(output_dir, temp_name)
    print(f"Output: {output_filepath}")
    fullfile_xlwings.save(output_filepath)
    fullfile_xlwings.close()