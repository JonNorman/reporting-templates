#coding: utf-8
from __future__ import unicode_literals
import logging
import numpy as np
import re
from common import *

def drop_row_with_value_in_column(df, column_name, value, exact_match):

    if exact_match:
        filtered = df.groupby(df[column_name] == value)
    else:
        filtered = df.groupby(df[column_name].str.contains(value))

    if False not in filtered.groups:
        logging.error('All rows have been filtered out where "{}" {} value "{}". Ignoring cleaning step.'.format(
            column_name,
            'equals' if exact_match else 'contains',
            value))
        return df

    if True not in filtered.groups:
        logging.warning('No rows matching the criteria of "{}" with value {} "{}". Ignoring cleaning step.'.format(
            column_name,
            'equal to' if exact_match else 'containing',
            value))
        return df

    keep_rows, drop_rows = filtered.get_group(False), filtered.get_group(True)

    logging.info('Removing the following row(s) where "{}" {} value "{}":'.format(
        column_name,
        'equals' if exact_match else 'contains',
        value))

    log_divider()
    for index, row in drop_rows.iterrows():
        logging.info('{:>2} {}'.format(index, printable(row)))
    log_divider()

    return keep_rows

def drop_column(df, column_name):
    logging.debug('Removing "{0}" column.'.format(column_name))
    return df.drop(column_name, 1)

def replace_goal_quantity_with_nan(df, column_name, threshold):

    singular_goal_quantities = df[df[column_name] <= threshold]
    if not singular_goal_quantities.empty:
        logging.info('Altering the following lines'' goal quantities to "N/A"')
        log_divider()
        for index, row in singular_goal_quantities.iterrows():
            logging.info('{:>2} {}'.format(index, printable(row)))
        log_divider()

    df[df[[column_name]] <= threshold] = np.nan
    return df
