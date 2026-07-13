from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.styles import Alignment
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
import shutil

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
    temp_name = temp_name.replace("_41.", ".")
    final_filepath = os.path.join(final_dir, temp_name)
    print(f"Final (openpyxl): {final_filepath}")
    shutil.copy(output_filepath, final_filepath)