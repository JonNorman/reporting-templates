#coding: utf-8
from __future__ import unicode_literals
import logging

def column_names(df):

    column_names = ['Line Item',
                   'Creative size',
                   'Line item ID',
                   'Line item start date',
                   'Line item end date',
                   'Goal quantity',
                   'Delivery indicator',
                   'Ad server impressions',
                   'Ad server clicks',
                   'Ad server average eCPM (£)',
                   'Ad server CTR',
                   'Ad server CPM and CPC revenue (£)'
                   ]

    headers = list(df.columns.values)

    logging.debug('Validating column names; expecting following in data: ')
    logging.debug('\t' + '\n\t'.join(column_names))

    mismatches = [ (column, header) for column, header in zip(column_names, headers) if column != header]
    for column, header in mismatches:
        logging.warn('\tExpected "{0}", received "{1}"'.format(column, header))

    return len(mismatches) == 0 and len(column_names) == len(headers)
