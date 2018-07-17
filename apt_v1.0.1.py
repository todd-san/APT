# core python imports
import os
import sys
import time
import copy
import json
import logging
import datetime
import argparse

# 3rd party imports
import xlwt
import openpyxl
from openpyxl.chart import (
    ScatterChart,
    Reference,
    Series,

)
from scipy import stats


# relative imports
from src import apt_logger
from src import read_xlsx as rx
from src import AptSpec as apt
from src import write_xlsx as wx


def main():
    args = parse_args()
    log = apt_logger.init(
        debug=args.debug,
        file=args.debug_file,
        stream=args.debug_stream,
        full=args.debug_all
    )
    log.debug('ARGS: {}'.format(args))

    apt_specs = None
    order = None
    data = None

    if not args.csv and not args.xlsx:
        log.error('No input file defined')
        sys.exit(-1)

    if args.csv:
        log.debug('CSV')
        log.debug('args.csv: {}'.format(args.csv[0]))
        log.debug('args.start: {}'.format(args.start))
        log.debug('args.end: {}'.format(args.end))

    elif args.xlsx:
        if args.xlsx.split('.')[-1] == 'xlsx':
            log.info('reading [xlsx] from path: "{}"'.format(args.xlsx))
            log.debug('args.xlsx: {}'.format(args.xlsx[0]))
            log.debug('args.start: {}'.format(args.start))
            log.debug('args.end: {}'.format(args.end))
            if args.end:
                data, order = rx.read_xlsx(file=args.xlsx, start=int(args.start), log=log, end=int(args.end))
            else:
                data, order = rx.read_xlsx(file=args.xlsx, start=int(args.start), log=log)
        else:
            log.error('Wrong file type, given: "{}", expected: "{}"'.format(args.xlsx.split('.')[-1], 'xlsx'))
            sys.exit(-1)

        log.info('finished reading')
        log.debug('spec order: {}'.format(order))
        apt_specs = classify_data(data, log)

    if args.trim:
        log.debug('TRIM')

    if apt_specs and order:
        log.debug('-'*75)
        # log.debug('apt_specs: {}'.format(apt_specs))
        log.debug('-'*75)
        wx.write_xlsx(apt_specs, order, data, log)


def classify_data(data, log):
    log.info('-'*75)
    specs = []
    for i, key in enumerate(dict.keys(data)):
        specs.append(apt.AptSpec(key, **data[key]))

    for spec in specs:
        cf = spec.curve_fit()
        popt = cf['popt']
        pcov = cf['pcov']
        if not all([v == 0 for v in popt]):
            log.info('AptSpec: "{}" curve_fit successful'.format(spec.name))
            log.info('  popt: {}, {}, {}, curve_fit successful'.format(popt[0], popt[1], popt[2]))
            log.info('  pcov: {}, {}, {}, curve_fit successful'.format(pcov[0][0], pcov[1][1], pcov[2][2]))
        else:
            log.info('AptSpec: "{}" curve_fit failed'.format(spec.name))
    log.info('-' * 75)

    return specs


def parse_args():
    """
    :return: <Namespace> args.parse_args(), all argument key:value pairs as specified by function.
    """
    trim_item = 1
    debug_mode = 1
    analyze = 1

    trim_help = 'Toggle the trim function on. ' \
                '\n Specify WHAT with [-N] names, or [-S] samples,' \
                ' and WHERE with [-h] hours or [-n] number of samples'
    debug_help = 'Toggle debug to ON'

    analyze_help = 'Specify what to analyze within the test results'

    f_help = 'f: Toggle debug to file-mode only'
    c_help = 'c: Toggle debug to console/stream-mode only'
    a_help = 'a: Toggle debug mode to all'
    s_help = 's: What test sample(s) to use for the [--trim] function'
    l_help = 'l: Labels for the [--trim] function'
    t_help = 't: Time(hours) for WHERE to [--trim]'
    n_help = 'n: n samples for WHERE to [--trim]'
    xlsx_help = 'xlsx: raw *.xlsx test file'
    csv_help = 'csv: raw *.csv test file'
    start_help = 'start: Where to start analyzing test data'
    end_help = 'end: Where to end analyzing test data'

    args = argparse.ArgumentParser()

    args.add_argument('--trim', const=trim_item, action='store_const', dest='trim', default=0, help=trim_help)
    args.add_argument('--debug', const=debug_mode, action='store_const', dest='debug', default=0, help=debug_help)
    args.add_argument('--analyze', const=analyze, action='store_const', dest='analyze', default=0, help=analyze_help)
    args.add_argument('-f', const=1, action='store_const', dest='debug_file', default=0, help=f_help)
    args.add_argument('-c', const=1, action='store_const', dest='debug_stream', default=0, help=c_help)
    args.add_argument('-a', const=1, action='store_const', dest='debug_all', default=0, help=a_help)
    args.add_argument('-s', type=int, nargs='+', help=s_help)
    args.add_argument('-l', type=str, nargs='+', help=l_help)
    args.add_argument('-t', type=str, nargs='+', help=t_help)
    args.add_argument('-n', type=str, nargs='+', help=n_help)
    args.add_argument('-xlsx', type=str, help=xlsx_help)
    args.add_argument('-csv', type=str, nargs='+', help=csv_help)
    args.add_argument('-start', type=int, help=start_help, default=0)
    args.add_argument('-end', type=int, help=end_help, default=None)

    return args.parse_args()


if __name__ == '__main__':
    main()
