from __future__ import unicode_literals

from functools import reduce
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
        return df

    keep_rows, drop_rows = filtered.get_group(False), filtered.get_group(True)

    logging.info('Removing the following row(s) where "{}" {} "{}":'.format(
        column_name,
        'equals' if exact_match else 'contains',
        value))

    for index, row in drop_rows.iterrows():
        logging.info('\t{:>2} {}'.format(index, printable(row)))

    return keep_rows

def replace_column_value(df, column_name, pattern, replacement):
    logging.debug('Replacing "{}" with "{}" in column "{}"'.format(pattern, replacement, column_name))
    df.loc[df[column_name] == pattern, column_name] = replacement
    return df

def replace_column_extract(df, column_name, extract_value_from_row):
    logging.debug('Replacing "{}" with extracted value by extract.'.format(column_name))
    df[column_name] = df.apply (lambda row: extract_value_from_row(row), axis=1)
    return df

def drop_columns_not_required(df, required_columns):

    actual_columns = list(df.columns.values)
    ignored_columns = [ column for column in actual_columns if column not in required_columns ]

    logging.debug('Removing columns: "{0}" column.'.format(', '.join(ignored_columns)))
    return reduce(lambda df, column: df.drop(column, 1), ignored_columns, df)

def replace_value_below_threshold_with_nan(df, column_name, threshold):

    singular_goal_quantities = df[df[column_name] <= threshold]
    if not singular_goal_quantities.empty:
        logging.info('Overwriting {} values under {} with "N/A"'.format(column_name, threshold))
        for index, row in singular_goal_quantities.iterrows():
            logging.info('\t{:>2} {}'.format(index, printable(row)))

    df.loc[df[column_name] <= threshold, column_name] = np.nan
    return df

def replace_datetime_with_date(df, column_name):
    logging.debug('Converting {} from datetime to date.'.format(column_name))
    df[column_name] = df[column_name].dt.date if 'Unlimited' not in df[column_name].values else df[column_name]
    return df

def reorder_columns(df, ordered_columns):
    logging.debug('Reordering dataframe to have columns: {}.'.format(','.join(ordered_columns)))
    original_columns = set(df.columns.tolist())
    updated_columns = set(ordered_columns)
    if original_columns == updated_columns:
        return df[ordered_columns]
    else:
        logging.error('Columns provided for reordering do not match those already present in the dataframe.')
        logging.error('Columns in one and not the other: {}.'.format(','.join(original_columns ^ updated_columns)))
        return df
