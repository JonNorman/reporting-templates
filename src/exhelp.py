"""Excel helper methods powered by openpyxl"""

# Python stdlib imports
from __future__ import unicode_literals
from collections import defaultdict
from itertools import product
import logging
import re

# package imports
from common import *

# openpyxl imports
from openpyxl.styles import Font
from openpyxl.styles.borders import Border, Side

def get_cells_by_regex(worksheet, pattern):
    matcher = re.compile(pattern)
    return [ c for c in worksheet.get_cell_collection() if matcher.match(str(c.value)) ]

def get_cell_by_regex(worksheet, pattern):
    all_matches = get_cells_by_regex(worksheet, pattern)
    return (all_matches[0] if all_matches else None)

def range_size(first_cell, last_cell):
    """ Returns the height and width of the range `start_cell`:`end_cell`

        :param first_cell: the top left cell in the range
        :type first_cell: Cell

        :param last_cell: the bottom right cell in the range
        :type last_cell: Cell

        :rtype: Tuple(height: int, width: int)
    """
    return (last_cell.row - first_cell.row + 1, last_cell.col_idx - first_cell.col_idx + 1)

def cells_between(first_cell, last_cell):
    """ Returns a list of the cells between first_cell and last_cell (inclusive)

        :param first_cell: the top left cell in the range
        :type first_cell: Cell

        :param last_cell: the bottom right cell in the range
        :type last_cell: Cell
    """
    height, width = range_size(first_cell, last_cell)
    return [ first_cell.offset(row = row, column = column) for row, column in product(range(height), range(width) )]

def get_range(from_first, to_last):
    return '{}:{}'.format(from_first.coordinate, to_last.coordinate)

def get_sum_formula(from_first, to_last):
    return '=SUM({})'.format(get_range(from_first, to_last))

def get_add_formula(cells):
    """ Returns a formula calculating the addition of each cell in cells

        :param cells: cells that need to be summed
        :type cells: List[Cell]
    """
    return '={}'.format('+'.join([ cell.coordinate for cell in cells ]))

##########################
######Style Helpers#######
##########################

def get_border(top_left, bottom_right, cell, style = 'medium'):
    """Given a range, return the borders of a cell such that all the cells on
    the edge have the appropriate side borders. If a cell is not on the
    perimeter of the range then there should be no borders.
    """
    border = Side(style = style)
    borders = {
        'left': border if cell.col_idx == top_left.col_idx else Side(None),
        'right': border if cell.col_idx == bottom_right.col_idx else Side(None),
        'top': border if cell.row == top_left.row else Side(None),
        'bottom': border if cell.row == bottom_right.row else Side(None)
    }
    return Border(**borders)


def update_font(template_font, updates):
    """Return a new font that copies the properties of `template_font` with
    any ovveriding properties taken from updates, a dict of prop -> value.
    """
    return Font(name = updates.get('name', template_font.name),
                size = updates.get('size', template_font.size),
                bold = updates.get('bold', template_font.bold),
                italic = updates.get('italic', template_font.italic),
                vertAlign = updates.get('vertAlign', template_font.vertAlign),
                underline = updates.get('underline', template_font.underline),
                strike = updates.get('strike', template_font.strike),
                color = updates.get('color', template_font.color))
