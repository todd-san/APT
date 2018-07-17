import datetime
import xlrd


def read_xlsx(file, start, log, end=None):
    # starting read_xlsx() function debug.log
    log.debug('-'*75)
    log.debug('reading file: {}'.format(file))
    log.debug('  test start: {}'.format(start))

    # fix data offset & vars
    start = start + 3
    if end:
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
    return format_data_dict(specs, params, data, log)


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
        return [datetime.datetime(*xlrd.xldate_as_tuple(val, b.datemode)) for val in w.col_values(i, start)]


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
    else:
        data[key.lower()] = []
        temp = w.col_values(i, start)
        for item in temp:
            if item:
                data[key.lower()].append(item)

    return data


def format_data_dict(specs, params, data, log):
    """
    :param log: <Logger> transporting python logger into this function for debugging.
    :param specs: <list> of spec keys in data dict.
    :param params: <list> of temp and baro keys in data dict
    :param data: <dict> flat dictionary of the populated test data
    :return: d: <dict> dictionary of re-formatted test data
    :return: specs: unchanged from input.. returned to keep track of order
    """
    d = {}
    for spec in specs:
        d[spec] = {
            'psi': data[spec],
            'datetime': data['datetime']
        }
        for param in params:
            d[spec][param] = data[param]

    return d, specs
