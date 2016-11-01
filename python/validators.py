from __future__ import unicode_literals
import logging

# there will be columns that we don't want and will drop, but ensure that
# the columns we DO have are in the same order as in the data
def column_names(df, required_columns):

    actual_columns = list(df.columns.values)

    logging.debug('Validating column names; expecting following in data: ')
    logging.debug('\t' + '\n\t'.join(required_columns))

    # look for columns that we will ignore
    ignored = [ column for column in actual_columns if column not in required_columns ]
    kept =    [ column for column in actual_columns if column in required_columns ]

    for column in ignored:
        logging.warn('Found column "{}" in data; this column is not required and will be ignored.'.format(column))

    mismatches = [ (a, b) for a, b in zip(required_columns, kept) if a != b ]

    for column, header in mismatches:
        logging.warn('\tExpected "{0}", received "{1}"'.format(column, header))

    return len(mismatches) == 0
