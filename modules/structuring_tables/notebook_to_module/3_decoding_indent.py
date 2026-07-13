from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter 
import xlwings as xw
import re
import os
import copy
import sys
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.append(parent_dir)
from algorithm import get_consecutive_ranges
from getcolwidth import get_real_column_width

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