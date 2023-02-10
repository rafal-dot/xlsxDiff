# -*- coding: utf-8 -*-
"""
xlsxDiff.py: Excel .xlsx spreadsheet files comparison tool. It compares spreadsheets cell by cell and
produces output in Word changes tracking style on cell level

https://github.com/rafal-dot/xlsxDiff

Created on Sat Apr  9 19:26:10 2020

@author: Rafal Czeczotka <rafal dot czeczotka at gmail.com>

Copyright (c) 2020-2023 Rafal Czeczotka

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>
"""

import argparse

from difflib import SequenceMatcher

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

import xlsxwriter


def log_print_message(is_log_enabled, message):
    """
    Print log message on screen
    :param is_log_enabled: defines, is message should be logged
    :param message: message text
    :return: current message length
    """
    if is_log_enabled:
        trim_len = log_print_message.log_msg_len - len(message) if len(message) < log_print_message.log_msg_len else 0
        print(message + " " * trim_len + "\r", end="")
        log_print_message.log_msg_len = len(message)


def get_format(column, row, modified_rows_set, modified_columns_set, f_do_nothing, f_highlight,
               do_update_sets=False):
    """
    Updates sets of columns and rows, checks if cell should be highlighted and returns relevant format
    :param column: current column
    :param row: current row
    :param modified_rows_set: set of modified rows
    :param modified_columns_set: set of modified columns
    :param f_do_nothing: format to be returned, if cell should not be highlighted
    :param f_highlight: format to be returned, if cell should be highlighted
    :param do_update_sets: information, if relevant column and row should be added as modified
    :return: relevant format of cell
    """
    if do_update_sets:
        modified_rows_set.add(row)
        modified_columns_set.add(column)
    if (column == 0 and row in modified_rows) or (row == 0 and column in modified_cols):
        return f_highlight
    return f_do_nothing


parser = argparse.ArgumentParser(description="Compares two .xlsx spreadsheets, cell content by cell content. "
                                             "Creates an output spreadsheet with all detailed text diffences "
                                             "highlighted, in word procesor-like format, tracking changes in "
                                             "each cell. Data can be filtered by color: (i) a white "
                                             "backround indicates changed cell, (ii) gray backround indicates "
                                             "a changed cell, (iii) green backround in first row/column "
                                             "highlights changed columns/rows, (iv) underlined blue text indicates "
                                             "added content and (v) crossed out red text indicates removed text.",
                                 epilog="Copyright (C) 2020-2023 Rafal Czeczotka. "
                                        "This program comes with ABSOLUTELY NO WARRANTY. "
                                        "This is free software, and you are welcome to redistribute it "
                                        "under GNU GPL v3 conditions (see https://www.gnu.org/licenses/).")
parser.add_argument("input1", help="first input spreadsheet file to compare", type=str)
parser.add_argument("input2", help="second input spreadsheet file to compare", type=str)
parser.add_argument("output", help="output spreadsheet file with highlighted differences", type=str)
parser.add_argument("-f", "--formula", help="compare formula text in cells instead of data values (default: disabled)",
                    action="store_true")
parser.add_argument("-x", "--highlight", help="in each row, if there are any cell changed, mark the first cell "
                                              "with a green background (row 1:1). In each column, if there are any "
                                              "cell changed, mark the first cell with a green background "
                                              "(column A:A). This option is useful for large spreadsheets, to "
                                              "facilitate changes identification. To filter out chaged data, "
                                              "autofilter with color filter might be applied later in spreadsheet "
                                              "software (default: disabled)",
                    action="store_true")
parser.add_argument("-a", "--autofilter", help="add autofilter for changed tabs (default: disabled)",
                    action="store_true")
parser.add_argument("-e", "--noempty", help="ignore empty cells (default: disabled)", action="store_true")
parser.add_argument("-v", "--verbose", help="verbose output. As it takes time to process large spreadsheets, "
                                            "this option facilitates progress tracking", action="store_true")
parser.add_argument("-q", "--quiet", help="no output messages", action="store_true")
parser.add_argument("--version", action="version", version='%(prog)s 1.1.0 (2023-02-11)')
args = parser.parse_args()

log_print_message.log_msg_len = 0
log_print_message(not args.quiet, f"Loading spreadsheet: {args.input1}")
i1_wb = load_workbook(filename=args.input1, data_only=not args.formula)
log_print_message(not args.quiet, f"Loading spreadsheet: {args.input2}")
i2_wb = load_workbook(filename=args.input2, data_only=not args.formula)

tab_names_dict = {}
for i2_tab in i2_wb.sheetnames:  # First, tabs from latest spreadsheet ;-)
    tab_names_dict[i2_tab] = {2, }
for i1_tab in i1_wb.sheetnames:
    if i1_tab in tab_names_dict.keys():
        tab_names_dict[i1_tab].add(1)
    else:
        tab_names_dict[i1_tab] = {1, }

o_wb = xlsxwriter.Workbook(args.output)

# If your Python interpreter raises issue here, just replace text:
# "**f_common"
# with following text:
# ""text_wrap": True, "border": 1, "align": "top""
# Merging dictionaries is not 100% standardised between various versions of Python yet
f_common = {"text_wrap": True, "border": 1, "align": "top"}
f_simple = o_wb.add_format({**f_common})
f_added = o_wb.add_format({**f_common, "bold": True, "underline": True, "color": "blue"})
f_added_cell_modified = o_wb.add_format(
    {**f_common, "bold": True, "underline": True, "color": "blue", "bg_color": "#a9d171"})  # light green
f_removed = o_wb.add_format({**f_common, "bold": True, "font_strikeout": True, "color": "red"})
f_removed_cell_modified = o_wb.add_format(
    {**f_common, "bold": True, "font_strikeout": True, "color": "red", "bg_color": "#a9d171"})  # light green
f_modified_cell = o_wb.add_format({**f_common, "bg_color": "#a9d171"})  # light green
f_equal_cell = o_wb.add_format({**f_common, "bg_color": "#c0c0c0"})
f_equal_cell_modified = o_wb.add_format({**f_common, "bg_color": "#a9d171"})

for current_tab_key in tab_names_dict.keys():
    if tab_names_dict[current_tab_key] == {1, 2}:
        # Tab name exists in both spreadsheets, individual cells are compared

        o_ws = o_wb.add_worksheet(current_tab_key)
        i1_ws = i1_wb[current_tab_key]
        i2_ws = i2_wb[current_tab_key]
        is_data_in_tab_modified = False

        modified_rows = set()
        modified_cols = set()
        for c in range(max(i1_ws.max_column, i2_ws.max_column) - 1, -1, -1):
            log_print_message(not args.quiet, f"Comparing tab: {current_tab_key}  column: {c}")
            o_ws.set_column(c, c, max(
                i1_ws.column_dimensions[get_column_letter(c + 1)].width,
                i2_ws.column_dimensions[get_column_letter(c + 1)].width,
                2))  # 2 is for very narrow columns, to make them easier to spot ;-)

            for r in range(max(i1_ws.max_row, i2_ws.max_row) - 1, -1, -1):
                log_print_message(not args.quiet and args.verbose,
                                  f"Comparing tab: {current_tab_key}  column: {c}  row:{r:5}")
                (i1_value, i2_value) = (i1_ws.cell(r + 1, c + 1).value, i2_ws.cell(r + 1, c + 1).value)

                # Check if cells are (i) empty, (ii) new or (iii) removed
                if i1_value in {None, ""} and i2_value in {None, ""}:
                    # Both cells empty
                    if args.noempty:
                        continue
                    o_ws.write(r, c, i2_value,
                               get_format(c, r, modified_rows, modified_cols, f_equal_cell, f_equal_cell_modified))
                    continue
                elif str(i1_value) == str(i2_value):
                    # Both cells equal (but not empty)
                    o_ws.write(r, c, str(i1_value),
                               get_format(c, r, modified_rows, modified_cols, f_equal_cell, f_equal_cell_modified))
                    continue
                elif i1_value in {None, ""}:
                    # First cell empty (value added)
                    val = str(i2_value) if i2_value is not None else ""
                    o_ws.write(r, c, val,
                               get_format(c, r, modified_rows, modified_cols, f_added, f_added_cell_modified,
                                          do_update_sets=args.highlight))
                    is_data_in_tab_modified = True
                    continue
                elif i2_value in {None, ""}:
                    # Second cell empty (value removed)
                    val = str(i1_value) if i1_value is not None else ""
                    o_ws.write(r, c, val,
                               get_format(c, r, modified_rows, modified_cols, f_removed, f_removed_cell_modified,
                                          do_update_sets=args.highlight))
                    is_data_in_tab_modified = True
                    continue

                # None of 4 conditione above applicable, so compare content in detail
                rich_text = []
                is_output_rich_string = False
                for tag, i1, i2, j1, j2 in SequenceMatcher(None, str(i1_value), str(i2_value)).get_opcodes():
                    if tag == "equal":
                        rich_text.append(str(i1_value)[i1:i2])
                    elif tag == "delete":
                        rich_text.extend([f_removed, str(i1_value)[i1:i2]])
                        is_output_rich_string = True
                    elif tag == "insert":
                        rich_text.extend([f_added, str(i2_value)[j1:j2]])
                        is_output_rich_string = True
                    elif tag == "replace":
                        rich_text.extend([f_removed, str(i1_value)[i1:i2], f_added, str(i2_value)[j1:j2]])
                        is_output_rich_string = True

                if is_output_rich_string:
                    o_ws.write_rich_string(r, c, *rich_text,
                                           get_format(c, r, modified_rows, modified_cols, f_simple, f_modified_cell,
                                                      do_update_sets=args.highlight))
                else:
                    # After 4 obvious checks in the beginning, this should not happen, but if...
                    o_ws.write(r, c, str(rich_text),
                               get_format(c, r, modified_rows, modified_cols, f_simple, f_modified_cell,
                                          do_update_sets=args.highlight))
                is_data_in_tab_modified = True

        if not is_data_in_tab_modified:
            o_ws.set_tab_color("gray")
        elif args.autofilter and args.highlight:
            o_ws.autofilter(0, 0, 0, max(i1_ws.max_column, i2_ws.max_column) - 1)

    elif tab_names_dict[current_tab_key] == {2, }:
        # Tab name exists only in 2nd spreadsheet, i.e. the tab is new, so no comaprison needed, just copy data
        # from 2nd input spreadsheet to output and use "added" formatting

        o_ws = o_wb.add_worksheet(current_tab_key)
        o_ws.set_tab_color("blue")
        i2_ws = i2_wb[current_tab_key]

        for c in range(i2_ws.max_column):
            log_print_message(not args.quiet, f"New tab: {current_tab_key}  column: {c}")
            o_ws.set_column(c, c, i2_ws.column_dimensions[get_column_letter(c + 1)].width)
            for r in range(i2_ws.max_row):
                log_print_message(not args.quiet and args.verbose, f"New tab: {current_tab_key}  column: {c}  row: {r}")
                val = str(i2_ws.cell(r + 1, c + 1).value) if i2_ws.cell(r + 1, c + 1).value is not None else ""
                o_ws.write(r, c, val, f_added)

    elif tab_names_dict[current_tab_key] == {1, }:
        # Tab name exists only in 1st spreadsheet i.e. the tab is removed, so no comaprison needed, just copy data
        # from 1st input spreadsheet to output and use "removed" formatting

        o_ws = o_wb.add_worksheet(current_tab_key)
        o_ws.set_tab_color("red")
        i1_ws = i1_wb[current_tab_key]

        for c in range(i1_ws.max_column):
            log_print_message(not args.quiet, f"Removed tab: {current_tab_key}  column: {c}")
            o_ws.set_column(c, c, i1_ws.column_dimensions[get_column_letter(c + 1)].width)
            for r in range(i1_ws.max_row):
                log_print_message(not args.quiet and args.verbose,
                                  f"Removed tab: {current_tab_key}  column: {c}  row: {r}")
                val = str(i1_ws.cell(r + 1, c + 1).value) if i1_ws.cell(r + 1, c + 1).value is not None else ""
                o_ws.write(r, c, val, f_removed)

log_print_message(not args.quiet, f"Saving spreadsheet: {args.output}")
o_wb.close()
log_print_message(not args.quiet, " ")
