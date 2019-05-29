from __future__ import unicode_literals
from builtins import input

import logging
import os
import sys

############################################
##############   Monad stuff  ##############
############################################
def bind(x, f):
    if x is None:
        return None
    else:
        return f(x)

############################################
############## Input & Output ##############
############################################

def printable(items, is_data_row = True):
    return ('|' if is_data_row else '') + ', '.join([ str(c) for c in items ])

def get_logger(level = logging.INFO):
    loggers = {
        logging.DEBUG: logging.debug,
        logging.INFO: logging.info,
        logging.WARNING: logging.warning,
        logging.ERROR: logging.error,
        logging.CRITICAL: logging.critical
    }

    return loggers.get(level, logging.info)

def log_divider(level = logging.INFO, symbol = '-'):
    divider = symbol * 150
    get_logger()(divider)

# taken from http://code.activestate.com/recipes/577058/
def query_yes_no(question, default="yes"):
    """Ask a yes/no question via raw_input() and return their answer.

    "question" is a string that is presented to the user.
    "default" is the presumed answer if the user just hits <Enter>.
        It must be "yes" (the default), "no" or None (meaning
        an answer is required of the user).

    The "answer" return value is True for "yes" or False for "no".
    """
    valid = {"yes": True, "y": True, "ye": True,
             "no": False, "n": False}
    if default is None:
        prompt = " [y/n] "
    elif default == "yes":
        prompt = " [Y/n] "
    elif default == "no":
        prompt = " [y/N] "
    else:
        raise ValueError("invalid default answer: '%s'" % default)

    while True:
        sys.stdout.write(question + prompt)
        choice = input().lower()
        if default is not None and choice == '':
            return valid[default]
        elif choice in valid:
            return valid[choice]
        else:
            sys.stdout.write("Please respond with 'yes' or 'no' "
                             "(or 'y' or 'n').\n")

def offer_clean_exit(output_path):
    logging.error('Skipping report creation for report "{}".'.format(output_path))
    if query_yes_no('* Output file "{}" is being abandoned - want me to delete it?'.format(output_path)):
        logging.info('Deleting file {}...'.format(output_path))
        os.remove(output_path)
    return

def should_continue():
    return query_yes_no('Happy to continue?')
