from __future__ import unicode_literals
from collections import defaultdict
from common import *
import itertools
import logging
import openpyxl
import re

def first_non_blank_cell_above(cell):
    for row in range(cell.row):
        cell = cell.offset(row = -1)
        if cell.value is None:
            return cell.offset(row = 1)

def find_cells_with_regex(worksheet, pattern):
    return [ cell for cell in worksheet.get_cell_collection() if pattern.match(str(cell.value)) ]

def get_tags(worksheet, expected_tags, unique):
    all_tags = find_cells_with_regex(worksheet, re.compile('^<(\w+)>$'))
    filtered_tags = [ tag for tag in all_tags if tag.value in expected_tags ]

    # if tags only occur once, return tag: cell
    if unique:
        tags =  { str(cell.value): cell for cell in filtered_tags }
    # otherwise return tag: list[cell]
    else:
        tags = defaultdict(list)
        for tag in filtered_tags:
            tags[str(tag.value)].append(tag)

    missing_tags = [ tag for tag in expected_tags if tag not in tags ]

    if missing_tags:
        logging.error('Missing expected tags in the template file.')
        logging.error('Expected: "{}"'.format(', '.join(expected_tags)))
        logging.error('Actual: "{}"'.format(', '.join(tags)))
        logging.error('Missing: "{}"'.format(', '.join(missing_tags)))
        return None
    else:
        return tags

def distance(start_cell, end_cell):
    return (end_cell.row - start_cell.row + 1, end_cell.col_idx - start_cell.col_idx + 1)

def get_cell_range(first_cell, last_cell):
    rows_diff, cols_diff = distance(first_cell, last_cell)
    return [ first_cell.offset(row = row, column = col) for row, col in itertools.product(rows_diff, cols_diff)]

def write_row_between_tags(start_tag, end_tag, values = None):

    _, width = distance(start_tag, end_tag)

    # verify that we have the same number of cells as values to write (if provided)
    if values and width != len(values):
        logging.warning('Start and end tags ("{}", "{}") do not outline the same sized range as the values provided: "{}".'.format(
            start_tag.coordinate,
            end_tag.coordinate,
            values
        ))
        return (start_tag, end_tag)
    elif values:
        logging.debug('{:>2} {}'.format(
            start_tag.row,
            printable(values)
        ))
    else:
        values = [None] * distance(start_tag, end_tag)[1]

    # create a range of offsets to write across the row
    for offset, value in zip(range(width), values):
        start_tag.offset(column = offset).value = value

    return (start_tag.offset(row = 1), end_tag.offset(row = 1))
