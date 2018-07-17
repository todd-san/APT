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
from pytz import timezone
import numpy as np
import xlwt
import xlrd
import openpyxl
from openpyxl.chart import (
    ScatterChart,
    Reference,
    Series,

)
import matplotlib.pyplot as plt
from scipy import stats
from scipy.optimize import curve_fit


# relative imports
from src import apt_logger
from src import read_xlsx as rx
from src import AptSpec as apt


def main():
    args = parse_args()
    log = apt_logger.init(
        debug=args.debug,
        file=args.debug_file,
        stream=args.debug_stream,
        full=args.debug_all
    )

    log.info('ARGS: {}'.format(args))

    if not args.csv and not args.xlsx:
        log.error('No input file defined')
        sys.exit(-1)

    if args.csv:
        log.debug('CSV')
        log.debug('args.csv: {}'.format(args.csv[0]))
        log.debug('args.start: {}'.format(args.start))
        log.debug('args.end: {}'.format(args.end))

    elif args.xlsx:
        log.info('reading [xlsx] from path: "{}"'.format(args.xlsx))
        log.debug('args.xlsx: {}'.format(args.xlsx[0]))
        log.debug('args.start: {}'.format(args.start))
        log.debug('args.end: {}'.format(args.end))
        if args.end:
            data = rx.read_xlsx(file=args.xlsx, start=int(args.start), log=log, end=int(args.end))
        else:
            data = rx.read_xlsx(file=args.xlsx, start=int(args.start), log=log)

        log.info('finished reading')
        classify_apt_data(data, log)

    if args.trim:
        log.debug('TRIM')


def classify_apt_data(data, log):
    specs = []

    for i, key in enumerate(dict.keys(data)):
        if i == 0:
            log.info('AptSpec made: data[{}]'.format(key))
            log.info('dict.keys(data[{}])'.format(dict.keys(data[key])))
            specs.append(apt.AptSpec(key, **data[key]))

    log.info('spec.name: {}'.format(specs[0].name))
    log.info('spec.psi[:10]: {}'.format(specs[0].psi[:10]))
    log.info('spec.baro[:10]: {}'.format(specs[0].baro[:10]))
    log.info('spec.baro[:10]: {}'.format(specs[0].curve_fit))


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
