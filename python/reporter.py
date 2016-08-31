#coding: utf-8
from __future__ import unicode_literals
from cleaners import *
from exhelp import *
from validators import *
from openpyxl.utils.dataframe import dataframe_to_rows

from functools import partial
import logging
import openpyxl
import os
import pandas
import re
import shutil
import sys
import xlrd

logging.basicConfig(level = logging.DEBUG, format = '> %(message)s')

def read_inputs(inputs_dir = 'inputs',
                input_sheet_name = 'Report data',
                validators = [],
                cleaners = []):

    # read input file(s) from input directory into a raw df
    logging.info('Searching for excel input files...')

    def xl_inputs():
        # for each directory
        for (dirpath, dirnames, filenames) in os.walk(inputs_dir):
            for filename in filenames:
                if filename[:2] != '~$':
                    yield os.path.join(dirpath,filename)

    # for each file in the xl_filenames, validate and clean
    for xl_input in xl_inputs():

        logging.info('Reading in file "{0}"...'.format(xl_input))
        df_raw = pandas.read_excel(io = xl_input, sheetname = input_sheet_name)

        # if validators have been provided, use them
        logging.info('Validating data...')
        if validators:
            if all([ validator(df_raw) for validator in validators ]):
                logging.debug('Input data is valid according to provided validator(s).')
            else:
                logging.error('Input data is invalid according to provided validator;' +
                      'file {0} will be skipped.'.format(xl_input)
                      )
                continue

        # if cleaners have been provided, use them
        logging.info('Cleaning data...')
        if cleaners:
            df_clean = reduce(lambda df, cleaner: cleaner(df), cleaners, df_raw)
        else:
            df_clean = df_raw

        yield (df_clean, os.path.basename(xl_input))

def copy_template_to_outputs(template_path, output_path):
    logging.info('Copying template from "{0}" to "{1}"...'.format(template_path, output_path))

    if (os.path.isfile(output_path)):
        overwrite = query_yes_no('* File "{0}" already exists. Confirm overwrite?'.format(output_path))
        if overwrite:
            logging.info('Overwriting "{0}"...'.format(output_path))
        else:
            logging.info('Not overwriting file; ignoring this report...')
            return False
    else:
        logging.info('Creating report file  "{0}"...'.format(output_path))

    shutil.copyfile(template_path, output_path)
    return True

# Line items are grouped if they have the same start/end date and ORD-XXXXXXX-X-X number
def identify_line_item_groups(df):

    def extract_id(line_item_name):
        line_item_id_pattern = '.*?(ORD-\d+-\d+-\d+).*'
        matches = re.match(line_item_id_pattern, line_item_name)
        return matches.group(1)

    df['Line Item Group'] = df['Line Item'].map(extract_id)
    df['Line Item Group'] = df['Line Item Group'] + '-' + df['Line item start date'].map(str)
    df['Line Item Group'] = df['Line Item Group'] + '-' + df['Line item end date'].map(lambda x: str(x))
    return df

def write_data(df, output_path):

    logging.info('Writing report data to "' + output_path + '"...')

    # open up the workbook
    workbook = openpyxl.load_workbook(output_path)
    worksheet = workbook._sheets[0]

    # group alike line items together and prepare to write them to the worksheet
    df = identify_line_item_groups(df)
    line_item_groups = df \
                        .sort_values(by = 'Creative size') \
                        .groupby(by = ['Line Item Group'], sort = ['Line item end date', 'Line item start date'])

    line_item_groups = { group_id: drop_column(line_items, 'Line Item Group') for (group_id, line_items) in  line_item_groups }

    tags = find_tags(worksheet)
    logging.debug('Found the following tags in the worksheet: {}'.format(','.join(tags.keys())))

    expected_tags = ['<header_start>', '<header_end>', '<data_start>', '<data_end>', '<order_id>']
    missing_tags = [ tag for tag in expected_tags if tag not in tags ]

    if missing_tags:
        logging.error('Missing expected tags in the template file.')
        logging.error('Expected: "{}"'.format(', '.join(expected_tags)))
        logging.error('Actual: "{}"'.format(', '.join(tags)))
        logging.error('Missing: "{}"'.format(', '.join(missing_tags)))
        logging.error('Skipping report creation for report "{}".'.format(output_path))
        if query_yes_no('* Output file "{}" is being abandoned - want me to delete it?'.format(output_path)):
            logging.info('Deleting file {}...'.format(output_path))
            os.remove(output_path)
        return

    # replace the header row with the headers from the groups
    headers = list(line_item_groups.itervalues().next().columns.values)
    logging.debug('Writing headers to output...')
    write_row_between_tags(worksheet, tags['<header_start>'], tags['<header_end>'], headers)

    # write the data rows
    start_tag, end_tag = tags['<data_start>'], tags['<data_end>']
    for _, line_items in line_item_groups.iteritems():

        if len(line_items) > 1:
            start_tag, end_tag = write_blank_row_between_tags(worksheet, start_tag, end_tag)

        for row in dataframe_to_rows(line_items, index = False, header = False):
            start_tag, end_tag = write_row_between_tags(worksheet, start_tag, end_tag, row)

        # if we have more than one line item in a group, then we need to provide a subtotal row
        if len(line_items) > 1:
            total_row  = ['<total>', '<total_impressions>', '<total_clicks>', '<total_clickthrough>']
            start_tag, end_tag = write_total_row_between_tags(worksheet, start_tag, end_tag, total_row)

    # add a grand total row
    start_tag, end_tag = write_blank_row_between_tags(worksheet, start_tag, end_tag)
    total_row  = ['<grand_total>', '<grand_total_impressions>', '<grand_total_clicks>', '<grand_total_clickthrough>']
    start_tag, end_tag = write_total_row_between_tags(worksheet, start_tag, end_tag, total_row)

    # replace the order_id tag
    order_id = re.compile('.*?(ORD-\d+)-.*').match(df['Line Item'].iloc[0]).group(1)
    logging.info('Determined Order ID as "{}"; writing to output.'.format(order_id))
    worksheet[tags['<order_id>'].coordinate] = order_id

    workbook.save(output_path)

    return

def apply_styling(workbook):
    # borders
    # merged order_id
    # total rows
    # bottom grey row
    pass

def replace_tags(workbook):
    expected_tags = ['<header_start>', '<header_end>', '<data_start>', '<data_end>', '<order_id>']
    tags = find_tags(worksheet)



def main():
    cleaners = [partial(drop_row_with_value_in_column, column_name = 'Line Item',
                                                       value = 'Total',
                                                       exact_match = True),
                partial(drop_row_with_value_in_column, column_name ='Line Item',
                                                       value = 'TEST',
                                                       exact_match = False),
                partial(drop_column, column_name = 'Line item ID'),
                partial(drop_column, column_name = 'Ad server average eCPM (£)'),
                partial(drop_column, column_name = 'Ad server CPM and CPC revenue (£)'),
                partial(replace_goal_quantity_with_nan, column_name = 'Goal quantity',
                                                        threshold = 1000)]

    validators = [column_names]

    template_path = 'template.xlsx'

    # read in the xl_filenames, validate and clean
    dfs_and_name = read_inputs(validators = validators, cleaners = cleaners)

    for df, original_filename in dfs_and_name:
        output_path = 'outputs/' + original_filename

        # can make this nomadic - each time returning a workbook and the output path
        if copy_template_to_outputs(template_path = 'template.xlsx',
                                    output_path = output_path):
            write_data(df, output_path)

if __name__ == '__main__':
    main()
