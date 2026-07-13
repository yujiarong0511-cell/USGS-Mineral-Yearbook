from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter 
import xlwings as xw
import re
import os
import sys
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.append(parent_dir)
from algorithm import get_consecutive_ranges
from getcolwidth import get_real_column_width
import copy

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