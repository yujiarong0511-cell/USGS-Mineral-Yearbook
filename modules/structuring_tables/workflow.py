import copy
import os
import re
import shutil
import sys
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.append(parent_dir)

# Excel-processing packages
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter, column_index_from_string
import xlwings as xw

# Self-defined functions
from algorithm import get_consecutive_ranges
from getcolwidth import get_real_column_width

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


def cleansing_table(continent_select, country_year):
    # %%
    # continent_select = "southamerica" [INPUT VAR 1]
    input_dir = "../../data/structured_tables/temp/usgs_" + continent_select + "/" # Temp files
    output_dir = input_dir

    # %% [markdown]
    # ##### Take Single Country (2019) as an example (encapsulated into .py)

    # %%
    name = country_year + "_1.xlsx" # [INPUT VAR 2] # IMPORTANT: suffix "_1" is a must (indicative of output from the first script)
    temp_name = copy.copy(name) 

    # %% [markdown]
    # ##### 2.1 Restore entries spanning across several rows

    # %%
    input_filepath = os.path.join(input_dir, temp_name)
    fullfile_openpyxl = load_workbook(input_filepath)
    print(f"Input (openpyxl): {input_filepath}")
    table_openpyxl = list(fullfile_openpyxl)[2] # DEBUG: Table 2 are at times not named as 'Table 2' in Excel (e.g., in China-2019 as 'Table2')
    # Check layout
    if len(list(table_openpyxl.columns)) == 5: # Whether there happens to be 5 columns
        print("Layout check passed :)")
    else:
        print("Check layout before move on!")

    # %% [markdown]
    # **2.1.1 Identify blocks of split entries** (Criteria: horizontal border style)

    # %%
    ## Identify blocks of split entries by horizontal border style
    all_groups = [] # In this output list, each element corresponds to a complete entry
    current_group = []
    for row in table_openpyxl.iter_rows(min_col=1, max_col=1):
        cell = row[0]
        row_num = row[0].row
        top_border = cell.border.top.style
        bottom_border = cell.border.bottom.style
        if row_num != 1:
            # IMPORTANT DEBUG: Two-sided check of border style is crucial!
            previous_cell = table_openpyxl[row_num-1][0]
            previous_cell_top_border = previous_cell.border.top.style
            previous_cell_bottom_border = previous_cell.border.bottom.style
            if top_border is not None or previous_cell_bottom_border is not None: # Two-sided check of border style
                all_groups.append(current_group)
                current_group = []
            current_group.append(row_num)
        else:
            current_group.append(row_num)
    all_groups.append(current_group)

    print(all_groups)

    # %% [markdown]
    # **2.1.2 Concatenate contents to the first row in each entry block**

    # %%
    max_col = table_openpyxl.max_column
    for group in all_groups:
        if len(group) == 1:
            continue
        else:
            for col_index in range(1, max_col + 1): # In Excel col_index starts from 1
                current_col_first_row_cell = table_openpyxl.cell(row=group[0], column=col_index) # NOTE: We do not recommend use "table_openpyxl[`row_index`][`col_index-1`]" as while both row and column index starts from 1 in excel, column index SHOULD start from 0 in this command with dual square brackets (which could easily get us confused).
                concatenated_value = ""
                for row_index in group:
                    current_segment_cell = table_openpyxl.cell(row=row_index, column=col_index)
                    current_segment_cell_value = current_segment_cell.value
                    current_segment_cell_value_str = str(current_segment_cell_value).strip(" ") if current_segment_cell_value is not None else ""
                    concatenated_value = concatenated_value + ' ' + str(current_segment_cell_value_str)
                current_col_first_row_cell.value = concatenated_value.strip(" ")

    # Save modified table to new XLSX file
    temp_name = temp_name.replace("_1.", "_21.")
    output_filepath = os.path.join(output_dir, temp_name)
    print(f"Output (openpyxl): {output_filepath}")
    fullfile_openpyxl.save(output_filepath)

    # %% [markdown]
    # ##### 2.2 Identify and remove redundant rows (rows other than the first in each entry block)

    # %% [markdown]
    # **2.2.1 Identify rows other than the first in each entry block**

    # %%
    # Identify redundant rows
    all_groups.sort(reverse=True) # Sort outer layer items
    print(all_groups)

    [group.sort(reverse=True) for group in all_groups] # Sort inner layer items
    [group.pop() for group in all_groups] # Remove last item in each vested list in the inner layer
    print(all_groups)

    redundant_rows = []
    for group_rm in all_groups:
        redundant_rows = redundant_rows + group_rm # Concatenate all remaining items from vested lists
    print(redundant_rows)

    # Create concise strings readable to xlwings
    redundant_rows.sort(reverse=False) # Sort list of redundant rows in ascending order to fit the input of get_consecutive_ranges() function
    print(redundant_rows)
    redundant_rows_concise_list = get_consecutive_ranges(redundant_rows)
    print(redundant_rows_concise_list)
    # Slice list into groups (no more than 25 items each)
    slice_size = 25
    redundant_rows_concise_slices = [redundant_rows_concise_list[i:i+slice_size] for i in range(0, len(redundant_rows_concise_list), slice_size)]
    print(redundant_rows_concise_slices)
    # Concatenate items in each slice into a string
    redundant_rows_concise_str_segments = [','.join(sl) for sl in redundant_rows_concise_slices]
    print(redundant_rows_concise_str_segments)
    redundant_rows_concise_str_segments.reverse() # Reverse the order of slices to remove rows bottom-up
    print(redundant_rows_concise_str_segments)

    # %% [markdown]
    # **2.2.2 Conduct removal**

    # %%
    ## Operate removal
    # [IMPORTANT] Read-in half-processed file
    input_filepath = copy.copy(output_filepath)
    print(f"Input (xlwings): {input_filepath}")

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

    ## Remove all redundant rows by iterating row segments
    for row_segment in redundant_rows_concise_str_segments:
        table_xlwings.api.Range(row_segment).Delete()

    # [IMPORTANT] Overwrite half-processed file
    temp_name = temp_name.replace("_21.","_22.")
    output_filepath = os.path.join(output_dir, temp_name)
    print(f"Output (xlwings): {output_filepath}")
    fullfile_xlwings.save(output_filepath)
    fullfile_xlwings.close()

    # %% [markdown]
    # ##### 2.3* (Conditional) Merge split "Commodity" column and Remove the redundant column B

    # %%
    # [IMPORTANT] Read-in half-processed file
    input_filepath = copy.copy(output_filepath)
    print(f"Input (openpyxl): {input_filepath}")
    fullfile_openpyxl = load_workbook(input_filepath)
    table_openpyxl = list(fullfile_openpyxl)[2]
    # Check layout
    if len(list(table_openpyxl.columns)) == 5: # Whether there happens to be 5 columns
        print("Layout check passed :)")
    else:
        print("Check layout before move on!")

    # %%
    # IMPORTANT DEBUG: If there should be more than 5 columns while Col A happens to be much narrower than Col B ("Commodity" information spans across two columns), we need to merge Col A and B before moving on.
    # Note: Ad hoc for Indonesia-2019
    real_col_width_list = get_real_column_width(table_openpyxl) # IMPORTANT: table_openpyxl.column_dimensions['D'].width can return false value 13.0 (default value at the back end of openpyxl when apparent width is not attainable, given width of adjacent columns are sometimes compressed in excel). Thus we encapsulated a module to decompress and reveal the real width.
    # VERY IMPORTANT DEBUG: Column removal operation is only conducted when following condition is satisfied
    if len(list(table_openpyxl.columns)) > 5 and any([cell.value for cell in table_openpyxl['F']]) and real_col_width_list['A'] < real_col_width_list['B'] * 0.25: # DEBUG: (1) Real measurement of column width; (2) Checking if the rightmost redundant column is empty
        for row_index in range(1, table_openpyxl.max_row+1): # Iterate from the first row
            cell_A = table_openpyxl.cell(row=row_index, column=1)
            cell_B = table_openpyxl.cell(row=row_index, column=2)
            if cell_B.value is not None: # DEBUG: Replace "or" with "and". The determinator is equivalent to `not (cell_B.value is None or str(cell_B.value).strip() == "")`
                cell_A.value = cell_B.value
                new_aligngment = copy.copy(cell_B.alignment)
                new_aligngment.indent = 1.0
                cell_A.alignment = new_aligngment

        ### IMPORTANT DEBUG: The operation is only conducted when the above condition is satisfied
        ## 1. Save merging operation
        # [IMPORTANT] Overwrite half-processed file
        # IMPORTANT DEBUG: Here we do not rename the output file as it is indefinite whether the operation would be conducted (which would in turn cause failure of rename in subsequent blocks)
        output_filepath = os.path.join(output_dir, temp_name)
        print(f"Output (openpyxl): {output_filepath}")
        fullfile_openpyxl.save(output_filepath)

        ## 2. Delete redundant column B
        input_filepath =  copy.copy(output_filepath)
        print(f"Input (xlwings): {input_filepath}")
        # To save time, we initiate a new Excel process only if Excel is not running at the back end.
        if len(xw.apps) > 0:
            app = xw.apps.active
            print("Running Excel process is invoked :)")
        else:
            app = xw.App(visible=True)
            print("New Excel process has started!")

        # Open the original table
        fullfile_xlwings = app.books.open(input_filepath) # Open .xlsx file
        table_xlwings = fullfile_xlwings.sheets[2]
        ## Delete the redundant column B
        table_xlwings.api.Range("B:B").Delete()

        # Save the modified copy and close the original doc
        # [IMPORTANT] Overwrite half-processed file

        output_filepath = os.path.join(output_dir, temp_name)
        print(f"Output (xlwings): {output_filepath}")
        fullfile_xlwings.save(output_filepath)
        fullfile_xlwings.close()

    # %% [markdown]
    # ##### 2.4 Replace "Do."/"do." with clearly defined objects (value of closest upper neighbor)

    # %%
    # [IMPORTANT] Read-in half-processed file
    input_filepath = copy.copy(output_filepath)
    print(f"Input (openpyxl): {input_filepath}")
    fullfile_openpyxl = load_workbook(input_filepath)
    table_openpyxl = list(fullfile_openpyxl)[2]
    # Check layout
    if len(list(table_openpyxl.columns)) == 5: # Whether there happens to be 5 columns
        print("Layout check passed :)")
    else:
        print("Check layout before move on!")

    # %% [markdown]
    # **\*Additional check before substituting "Do./do.":** If after merging split rows there are still empty cells ("missing values") in the first column ("Commodity" column), set its value and layout as its closest upper neighbor. (e.g., Indonesia-2019)

    # %%
    # DEBUG: If after merging split rows there are still empty cells in the first column ("Commodity" column), set its value and layout as its closest upper neighbor.
    # Ad hoc for Indonesia-2019
    for row_index in range(2, table_openpyxl.max_row+1): # Iterate from the second row
        current_cell = table_openpyxl.cell(row=row_index, column=1)
        is_nonemptyrow = any([cell.value for cell in table_openpyxl[row_index]]) # DEBUG: Check if the row is NOT empty
        if current_cell.value is None and is_nonemptyrow: # DEBUG: Add a condition 'is_nonemptyrow' to make sure we only check rows in the body of the table (since at times max_row could embody empty rows beneath the body of the table)
            previous_cell = table_openpyxl.cell(row=row_index-1, column=1)
            current_cell.value = previous_cell.value # Replicate value
            if previous_cell.alignment is not None:
                    current_cell.alignment = copy.copy(previous_cell.alignment) # Replicate layout

    # %% [markdown]
    # Iterate through column A~E. In each column, match cells with "Do./do." with their reference cell and encapsulate them in a list where the first item is the reference cell. Then replace the alignment (indentation) and value of "Do./do." cells with their reference's.

    # %%
    # Iterate through column A~E. In each column, match cells with "Do./do." with their reference cell and encapsulate them in a list where the first item is the reference cell. Then replace the alignment (indentation) and value of "Do./do." cells with their reference's.
    for col_index in range(1, 5+1):
        for row_index in range(2, table_openpyxl.max_row+1): # Iterate from the second row
            current_cell = table_openpyxl.cell(row=row_index, column=col_index)
            current_cell_value = current_cell.value
            current_cell_value_str = str(current_cell_value).strip() if current_cell_value is not None else ""
            if current_cell_value_str in ['do.','Do.','Do']: # DEBUG: Add "Do" without dot ('.') to the list (e.g., Indonesia-2019)
                previous_cell = table_openpyxl.cell(row=row_index-1, column=col_index)
                current_cell.value = previous_cell.value # Replicate value
                if previous_cell.alignment is not None:
                    current_cell.alignment = copy.copy(previous_cell.alignment) # Replicate layout

    # %% [markdown]
    # **\*Additional check: Remove all "—" and "Continue(d)" from cells (but keep the rows)**

    # %%
    # Remove all "—" and "Continue(d)" from cells (but keep the rows)
    clean_regex = r'[—\s]*continue(d)?\s*$'
    for row in table_openpyxl.iter_rows(min_col=1, max_col=1):
        current_cell = row[0]
        if current_cell.value is not None:
            current_cell.value = re.sub(pattern=clean_regex, repl='', string=str(current_cell.value), flags=re.IGNORECASE).strip()

    # %%
    # Save modified table to new XLSX file
    temp_name = temp_name.replace("_22.","_23.")
    output_filepath = os.path.join(output_dir, temp_name)
    print(f"Output (openpyxl): {output_filepath}")
    fullfile_openpyxl.save(output_filepath)


def decoding_indent(continent_select, country_year):
    # %%
    # continent_select = "southamerica" # [INPUT VAR 1]
    input_dir = "../../data/structured_tables/temp/usgs_" + continent_select + "/"
    output_dir = input_dir

    # %% [markdown]
    # ##### Take Single Country (2019) as an example (encapsulated into .py)

    # %%
    name = country_year + "_23.xlsx" # [INPUT VAR 2] # IMPORTANT: suffix "_2X" is a must (indicative of output from the second script)
    temp_name = copy.copy(name) 

    # %% [markdown]
    # ##### 3.1 Decode indentation in the first column (labels of different tiers)

    # %%
    # [IMPORTANT: Load intermediate work file from /test directory]
    input_filepath = os.path.join(input_dir, temp_name)
    print(f"Input (openpyxl): {input_filepath}")
    fullfile_openpyxl = load_workbook(input_filepath)
    table_openpyxl = list(fullfile_openpyxl)[2] # DEBUG: Table 2 are at times not named as 'Table 2' in Excel (e.g., in China-2019 as 'Table2')
    # Check layout
    if len(list(table_openpyxl.columns)) == 5: # Whether there happens to be 5 columns
        print("Layout check passed :)")
    else:
        print("Check layout before move on!")

    # %%
    max_row = table_openpyxl.max_row
    label_stack = [] # Utilize stack data structure

    # Initialize (first row)
    first_cell = table_openpyxl.cell(row=1, column=1)
    first_cell_value = str(first_cell.value) if first_cell.value is not None else ""
    first_cell_indent = first_cell.alignment.indent if first_cell.alignment is not None else 0.0
    label_stack.append({'layer': 0, 
                        'indent': first_cell_indent, 
                        'value': first_cell_value})

    for row_index in range(2, max_row + 1): # Iterate from the second row
        current_cell = table_openpyxl.cell(row=row_index, column=1)
        current_cell_value = str(current_cell.value) if current_cell.value is not None else "" # DEBUG: Add default value (e.g., Brazil-2019)
        current_cell_indent = current_cell.alignment.indent if current_cell.alignment is not None else 0.0
        if current_cell_indent == label_stack[-1]['indent']:  # Sibling label
            label_stack[-1]['value'] = current_cell_value
        elif current_cell_indent > label_stack[-1]['indent']: # Descendant label
            new_label = {'layer': label_stack[-1]['layer']+1,
                        'indent': current_cell_indent,
                        'value': current_cell_value}
            label_stack.append(new_label)
        else: # current_cell_indent < label['indent']         # Ancestor label
            for label in reversed(label_stack): # Recursive
                if current_cell_indent == label['indent']:
                    label['value'] = current_cell_value
                    break
                else:
                    label_stack.pop()
        new_current_value = ' '.join([label['value'] for label in label_stack])
        current_cell.value = new_current_value
        new_current_alignment = copy.copy(current_cell.alignment)
        new_current_alignment.indent = 0.0
        current_cell.alignment = new_current_alignment

    # %%
    # Save modified table to new XLSX file
    temp_name = temp_name.replace("_23.", "_31.")
    output_filepath = os.path.join(output_dir, temp_name)
    print(f"Output (openpyxl): {output_filepath}")
    fullfile_openpyxl.save(output_filepath)

    # %% [markdown]
    # ##### 3.2 Remove redundant rows (where both "company" and "location" columns are empty)

    # %%
    # [IMPORTANT: Load intermediate work file `temp_name` from /test directory]
    input_filepath = copy.copy(output_filepath)
    print(f"Input (openpyxl): {input_filepath}")
    fullfile_openpyxl = load_workbook(input_filepath)
    table_openpyxl = list(fullfile_openpyxl)[2] # DEBUG: Table 2 are at times not named as 'Table 2' in Excel (e.g., in China-2019 as 'Table2')
    # Check layout
    if len(list(table_openpyxl.columns)) == 5: # Whether there happens to be 5 columns
        print("Layout check passed :)")
    else:
        print("Check layout before move on!")

    # %%
    redundant_rows = []
    for row_index in range(1, table_openpyxl.max_row+1): # Iterate from the first row
        cell_C_company = table_openpyxl.cell(row=row_index, column=3)
        cell_D_location = table_openpyxl.cell(row=row_index, column=4)
        if cell_C_company.value is None and cell_D_location.value is None:
            redundant_rows.append(row_index)

    if len(redundant_rows) > 0: # IMPORTANT DEBUG: Run the following block only if there SHOULD EXIST any redundant row.
        ## Create concise strings readable to xlwings (NOTE: Copied from `2-cleansing_table_single2019.ipynb`)
        redundant_rows.sort(reverse=False) # Sort list of redundant rows in ascending order to fit the input of get_consecutive_ranges() function
        print(redundant_rows)
        redundant_rows_concise_list = get_consecutive_ranges(redundant_rows)
        print(redundant_rows_concise_list)
        # Slice list into groups (no more than 25 items each)
        slice_size = 25
        redundant_rows_concise_slices = [redundant_rows_concise_list[i:i+slice_size] for i in range(0, len(redundant_rows_concise_list), slice_size)]
        print(redundant_rows_concise_slices)
        # Concatenate items in each slice into a string
        redundant_rows_concise_str_segments = [','.join(sl) for sl in redundant_rows_concise_slices]
        print(redundant_rows_concise_str_segments)
        redundant_rows_concise_str_segments.reverse() # Reverse the order of slices to remove rows bottom-up
        print(redundant_rows_concise_str_segments)

        ## Operate removal
        # [IMPORTANT] Read-in half-processed file
        input_filepath = copy.copy(output_filepath) # NOTE: THis input is exactly the same with the last (as they are intended to read the same table, though read-only and read-write respectively)
        print(f"Input: {input_filepath}")

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

        ## Remove all redundant rows by iterating row segments
        for row_segment in redundant_rows_concise_str_segments:
            table_xlwings.api.Range(row_segment).Delete()

        # Save modified table to new XLSX file
        output_filepath = os.path.join(output_dir, temp_name) # DEBUG: We do not rename the output table to be "_32" to avoid unnecessary complexity
        print(f"Output: {output_filepath}")
        fullfile_xlwings.save(output_filepath)
        fullfile_xlwings.close()


def filling_unit(continent_select, country_year):
    # %%
    # continent_select = "southamerica"
    origin_dir = "../../data/original_table/usgs_" + continent_select + "/"
    input_dir = "../../data/structured_tables/temp/usgs_" + continent_select + "/"
    output_dir = input_dir
    final_dir = "../../data/structured_tables/usgs_" + continent_select + "/"

    # %% [markdown]
    # ##### Take Single Country (2019) as an example (encapsulated into .py)

    # %%
    name = country_year + "_31.xlsx" # IMPORTANT: suffix "_3X" is a must (indicative of output from the 3rd script)
    origin_name = name.split('.')[0].split('_')[0] + "_" + name.split('.')[0].split('_')[1] + ".xlsx"
    temp_name = copy.copy(name)

    # %% [markdown]
    # ##### 4.1 Fill empty cells in "unit of measure" column with default value

    # %% [markdown]
    # **4.1.1 Identify default unit of measure**

    # %%
    # Retrieve default unit of measure from original table
    origin_filepath = os.path.join(origin_dir, origin_name)
    print(f"Origin (openpyxl): {origin_filepath}")
    fullfile_openpyxl = load_workbook(origin_filepath)
    table_openpyxl = list(fullfile_openpyxl)[2]

    # %%
    fuze = "unless otherwise specified"
    default_unit = None
    for col_index in range(1, table_openpyxl.max_column+1):
        for row_index in range(1, table_openpyxl.max_row+1):
            current_cell = table_openpyxl.cell(row=row_index, column=col_index)
            current_cell_value = current_cell.value # DEBUG: Discriminant is based on .value rather than the cell object itself.
            if current_cell_value is not None and fuze in str(current_cell_value).lower():
                default_unit = str(current_cell_value).strip().strip("()").replace(fuze, '').strip().strip(',').lower()
                break
        if default_unit is not None:
            break
    print(default_unit)

    # %% [markdown]
    # **4.1.2 Fill empty cells with default value**

    # %%
    # Read-in intermediate file
    input_filepath = os.path.join(input_dir, temp_name)
    print(f"Input (openpyxl): {input_filepath}")
    fullfile_openpyxl = load_workbook(input_filepath)
    table_openpyxl = list(fullfile_openpyxl)[2]
    # Check layout
    if len(list(table_openpyxl.columns)) == 5: # Whether there happens to be 5 columns
        print("Layout check passed :)")
    else:
        print("Check layout before move on!")

    # %%
    # Iterate through all cells in the second column (except for fully empty rows)
    # IMPORTANT: Only conduct the filling when "default_unit" is identified in the original table (or else skip the block)
    for row_index in range(1, table_openpyxl.max_row+1):
        cell_B = table_openpyxl.cell(row=row_index, column=2)
        is_nonemptyrow = any([cell.value for cell in table_openpyxl[row_index]])
        if cell_B.value is None and is_nonemptyrow:
            cell_B.value = default_unit if default_unit is not None else "(unspecified)"
            cell_B.alignment = Alignment(horizontal='right') # Prettify alignment by the way

    # Save modified table to new XLSX file
    temp_name = temp_name.replace("_31.", "_41.") # DEBUG: Replace "_32" with "_31"
    output_filepath = os.path.join(output_dir, temp_name)
    print(f"Output (openpyxl): {output_filepath}")
    fullfile_openpyxl.save(output_filepath)

    # %%
    # Copy the file to final directory
    temp_name = temp_name.replace("_41.", ".")
    final_filepath = os.path.join(final_dir, temp_name)
    print(f"Final (openpyxl): {final_filepath}")
    shutil.copy(output_filepath, final_filepath)