#coding: utf-8
from __future__ import unicode_literals
import xlutils

import logging
import openpyxl
import re

def find_cells_with_regex(worksheet, pattern):
    return [ cell for cell in worksheet.get_cell_collection() if pattern.match(str(cell.value)) ]

def find_tags(worksheet):
    return { str(cell.value): cell for cell in find_cells_with_regex(worksheet, re.compile('^<(\w+)>$')) }

def add_row_to_cell(worksheet, cell):
    return worksheet.cell(row = cell.row + 1, column = cell.col_idx)

def write_row_between_tags(worksheet, start_tag, end_tag, values):

    tag_range_size = end_tag.col_idx - start_tag.col_idx + 1
    if tag_range_size != len(values):
        logging.warning('Start and end tags ("{}", "{}") do not outline the same sized range as the values provided: "{}".'.format(
            start_tag.coordinate,
            end_tag.coordinate,
            values
        ))
        return (start_tag, end_tag)
    else:
        logging.debug('Writing values "{}" between start and end tags provided, "{}" and "{}"...'.format(
            values,
            start_tag.coordinate,
            end_tag.coordinate
        ))

        cols_with_values = zip(range(start_tag.col_idx, end_tag.col_idx + 1), values)

        for col_idx, value in cols_with_values:
            worksheet.cell(row = start_tag.row, column = col_idx).value = value

        return (add_row_to_cell(worksheet, start_tag), add_row_to_cell(worksheet, end_tag))

def write_blank_row_between_tags(worksheet, start_tag, end_tag):
    logging.debug('Writing blank values between start and end tags provided, "{}" and "{}"...'.format(
        start_tag.coordinate,
        end_tag.coordinate
    ))

    for col_idx in range(start_tag.col_idx, end_tag.col_idx + 1):
        worksheet.cell(row = start_tag.row, column = col_idx).value = ''

    return (add_row_to_cell(worksheet, start_tag), add_row_to_cell(worksheet, end_tag))

def write_total_row_between_tags(worksheet, start_tag, end_tag, values):
        logging.debug('Writing total row on line "{}"'.format(end_tag.row))

        cols_with_values = zip(range(end_tag.col_idx, end_tag.col_idx - len(values), -1), values[::-1])
        for col_idx, value in cols_with_values:
            worksheet.cell(row = end_tag.row, column = col_idx).value = value

        return (add_row_to_cell(worksheet, start_tag), add_row_to_cell(worksheet, end_tag))
