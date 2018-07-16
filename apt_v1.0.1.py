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


def main():
    args = parse_args()
    log = apt_logger.init(
        debug=args.debug,
        file=args.debug_file,
        stream=args.debug_stream,
        full=args.debug_all
    )

    if not args.csv and not args.xlsx:
        log.error('No input file defined')
        sys.exit(-1)

    if args.csv:
        log.debug('CSV')
        log.debug('args.csv: {}'.format(args.csv[0]))
        log.debug('args.start: {}'.format(args.start))
        log.debug('args.end: {}'.format(args.end))

    elif args.xlsx:
        log.debug('XLSX')
        log.debug('args.xlsx: {}'.format(args.xlsx[0]))
        log.debug('args.start: {}'.format(args.start))
        log.debug('args.end: {}'.format(args.end))
        data = read_xlsx(file=args.xlsx[0], start=int(args.start[0]), end=int(args.end[0]), log=log)
        apt=classify_apt_data(data=data, log=log)

    if args.trim:
        log.debug('TRIM')


def classify_apt_data(data, log):
    for key in dict.keys(data):
        for inner_key in dict.keys(data[key]):
            log.debug('{}:{} --> {}'.format(key, inner_key, type(data[key][inner_key])))


def read_xlsx(file, start, end, log):
    # starting read_xlsx() function debug.log
    log.debug('-'*75)
    log.debug('reading file: {}'.format(file))
    log.debug('  test start: {}'.format(start))

    # fix data offset & vars
    start = start + 3
    end = end + 3
    data = {}
    dates = []
    times = []
    specs = []
    params = []

    # open specified workbook
    wb = xlrd.open_workbook(file)

    # loop through the worksheets... this should probably be discarded, as there should only ever be 1
    for i in range(wb.nsheets):
        ws = wb.sheet_by_index(i)
        log.debug(" worksheet {}: {}".format(i + 1, ws.name))

        for j in range(ws.ncols):
            if j == 0:  # function read_date
                dates = read_date(wb, ws, j, start, end)

            elif j == 1:  # function read_time
                times = read_time(ws, j, start, end)

            else:
                if isinstance(ws.cell_value(1, j), (int, float)):  # read_spec_header
                    key = read_spec_header(ws, j)
                    specs.append(key)
                else:  # read_param_header
                    key = read_param_header(ws, j)
                    params.append(key)

                if key:  # read_data
                    data = read_data(ws, j, key, data, start, end)

        if dates and times:
            data['datetime'] = [datetime.datetime.combine(date, times[i]) for i, date in enumerate(dates)]

    # format data-object for easy class creation
    log.debug('-'*75)
    return format_data_dict(specs, params, data)


def format_data_dict(specs, params, data):
    """
    :param specs: <list> of spec keys in data dict.
    :param params: <list> of temp and baro keys in data dict
    :param data: <dict> flat dictionary of the populated test data
    :return: d: <dict> dictionary of re-formatted test data
    """
    d = {}
    for spec in specs:
        d[spec] = {
            'psi': data[spec],
            'time': data['datetime']
        }
        for param in params:
            d[spec][param] = data[param]

    return d


def read_spec_header(w, i):
    """
    :param w: <xlrd> worksheet
    :param i: <int> column #
    :return: <str> spec_header as str()
    """
    return (str(w.cell_value(0, i)) + '_s' + str(int(w.cell_value(1, i)))).lower()


def read_param_header(w, i):
    """
    :param w: <xlrd> worksheet (xlrd obj)
    :param i: <int> column #
    :return: <str> param_header as str()
    """
    return (str(w.cell_value(0, i)) + '_' + str(w.cell_value(1, i)))\
        .replace('[', '')\
        .replace(']', '')\
        .replace('°', 'deg_')\
        .replace('.', '_')\
        .strip('_')\
        .lower()


def read_time(w, i, start, end=None):
    """
    :param w: <xlrd> worksheet
    :param i: <int> denoting column #
    :param start: <int> denoting starting row #
    :param end: <int> denoting ending row #
    :return: <list> of datetime objects
    """
    if end:
        return [datetime.time((int(val * 24 * 3600) // 3600),       # hours
                              (int(val * 24 * 3600) % 3600) // 60,  # minutes
                              (int(val * 24 * 3600) % 60))          # seconds
                for val in w.col_values(i, start, end)]
    else:
        return [datetime.time((int(val * 24 * 3600) // 3600),  # hours
                              (int(val * 24 * 3600) % 3600) // 60,  # minutes
                              (int(val * 24 * 3600) % 60))  # seconds
                for val in w.col_values(i, start)]


def read_date(b, w, i, start, end=None):
    """
    :param b: <xlrd> workbook
    :param w: <xlrd> worksheet
    :param i: <int> denoting column #
    :param start: <int> denoting starting row #
    :param end: <int> denoted ending row #, default=None
    :return:
    """
    if end:
        return [datetime.datetime(*xlrd.xldate_as_tuple(val, b.datemode)) for val in w.col_values(i, start, end)]
    else:
        return [datetime.datetime(*xlrd.xldate_as_tuple(val, b.datemode)) for val in w.col_values(i, start, end)]


def read_data(w, i, key, data, start, end=None):
    """
    :param w: <xlrd> worksheet
    :param i: <int> denotes column # in worksheet
    :param key: <str> denotes str-object to use as dict key
    :param data: <dict> object to populate with data
    :param start: <int> denotes starting row #
    :param end: <int> denotes ending row #
    :return: data: <dict>
    """
    if end:
        data[key.lower()] = []
        temp = w.col_values(i, start, end)
        for item in temp:
            if item:
                data[key.lower()].append(item)
    return data


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
    args.add_argument('-xlsx', type=str, nargs='+', help=xlsx_help)
    args.add_argument('-csv', type=str, nargs='+', help=csv_help)
    args.add_argument('-start', type=str, nargs='+', help=start_help)
    args.add_argument('-end', type=str, nargs='+', help=end_help)

    return args.parse_args()


if __name__ == '__main__':
    main()
