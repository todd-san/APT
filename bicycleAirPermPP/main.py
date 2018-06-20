import xlrd
import numpy as np
import os
import xlwt
import json
import logging
import sys
import datetime
import time
from pytz import timezone
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit
from scipy import stats
from kapteyn import kmpfit

BLESSED_TEMP = 22.2222222  # deg C for normalization
CELSIUS_TO_KELVIN = 273.15


class AirPermTest:
    def __init__(self, **kwargs):
        self.specs = list()
        if kwargs:
            specs = []
            for key, value in kwargs.items():
                if key in ['datetime']:
                    setattr(self, key, value)
                elif key == 'baro':
                    self.pressure = value
                elif key[0] == 'p':
                    specs.append({
                        'name': key,
                        'pressures': value
                    })
                    # setattr(self, key, AirPermSpec(pressures=value))
                else:
                    setattr(self, 'temperatures', AirPermThermocouples(**{key: value}))
        else:
            log.error('No input **kwargs for AirPermTest Definition')
            raise TypeError

        for spec in specs:
            self.specs.append(AirPermSpec(name=spec['name'],
                                          time=self.time,
                                          pressures=spec['pressures'],
                                          temperatures=self.avg_temp_kelvin))

    @property
    def avg_temp_celsius(self):
        temps = list()
        if hasattr(self, 'temperatures'):
            for item in vars(self.temperatures):
                values = list()
                for value in getattr(self.temperatures, item):
                    values.append(value)
                temps.append(values)
            return list(np.mean(np.array(temps), axis=0))
        else:
            # given the way that properties seem to work, its okay if this passes through.
            pass

    @property
    def avg_temp_kelvin(self):
        if hasattr(self, 'temperatures'):
            temps = list()
            for item in vars(self.temperatures):
                values = list()
                for value in getattr(self.temperatures, item):
                    values.append(value+CELSIUS_TO_KELVIN)
                temps.append(values)
            return list(np.mean(np.array(temps), axis=0))
        else:
            pass

    @property
    def time(self):
        t0 = time.mktime(self.datetime[0].timetuple())

        t = [time.mktime(tN.timetuple()) - t0 for tN in self.datetime]

        return t


class AirPermSpec:
    def __init__(self, name, time, pressures, temperatures):

        self.name = name
        self.temps = temperatures                               # kelvin
        self.nom_p = pressures                                  # psi
        self.p = self.normalized_pressures(p_array=pressures)   # psi
        self.t = np.array([val/3600 for val in time])
        self.popt, self.pcov = curve_fit(self.exp_model, self.t, self.p)

    def normalized_pressures(self, p_array):
        norm_p = []
        for i, val in enumerate(p_array):
            norm_p.append((val*(BLESSED_TEMP+CELSIUS_TO_KELVIN))/(self.temps[i]))
        return np.array(norm_p)

    def exp_pb(self, conf=0.95):

        xd = self.t
        yd = self.p
        f_vars = self.popt
        x = np.linspace(xd.min(), xd.max(), 100)

        alpha = 1. - conf
        N = xd.size
        variables_n = len(f_vars)

        q = stats.t.ppf(1. - alpha / 2., N - variables_n)

        se = np.sqrt(1. / (N - variables_n) * np.sum((yd - self.exp_model(xd, *f_vars))**2))

        sx = (x - xd.mean()) ** 2
        sxd = np.sum((xd - xd.mean()) ** 2)

        yp = self.exp_model(x, *f_vars)

        dy = q * se * np.sqrt(1. + (1. / N) + (sx / sxd))

        lpb, upb = yp - dy, yp + dy

        return lpb, upb

    def exp_cb(self):
        x = np.linspace(self.t.min(), self.t.max(), 100)
        log.debug('VAR A: {}, 2*SIG_A: {}'.format(self.pcov[0][0], 2*self.pcov[0][0]**0.5))
        log.debug('VAR B: {}, 2*SIG_A: {}'.format(self.pcov[1][1], 2*self.pcov[1][1]**0.5))
        log.debug('VAR C: {}, 2*SIG_A: {}'.format(self.pcov[2][2], 2*self.pcov[1][1]**0.5))
        pos_a = self.popt[0]+(self.pcov[0][0]**0.5)
        neg_a = self.popt[0]-(self.pcov[0][0]**0.5)
        pos_b = self.popt[1]+(self.pcov[1][1]**0.5)
        neg_b = self.popt[1]-(self.pcov[1][1]**0.5)
        pos_c = self.popt[2]+(self.pcov[2][2]**0.5)
        neg_c = self.popt[2]-(self.pcov[2][2]**0.5)
        return self.exp_model(x, pos_a, neg_b, pos_c), self.exp_model(x, neg_a, pos_b, neg_c)

    def exp_kmpfit(self):
        # x = np.linspace(self.t.min(), self.t.max(), 100)
        x = self.t
        y = self.p
        f = kmpfit.simplefit(self.kmpfit_model, [.1, .1, .1], x, y)
        a, b, c = f.params
        dfdp = [np.exp(-b*x), a*x*np.exp(-b*x), 1]
        yhat, upper, lower = f.confidence_band(x, dfdp, 0.95, self.kmpfit_model)

        plt.scatter(x, y)
        ix = np.argsort(x)
        for i, l in enumerate((upper, lower, yhat)):
            plt.plot(x[ix], l[ix], c='g' if i == 2 else 'r', lw=2)
        plt.show()

    @staticmethod
    def exp_model(t, a, b, c):
        return a * np.exp(-b * t) + c

    @staticmethod
    def kmpfit_model(p, x):
        a, b, c = p
        return a*np.exp(-b*x) + c

    def exp_correlation(self):
        xd = self.t
        yd = self.p
        yp = self.exp_model(xd, *self.popt)
        n = np.size(yd)
        return (n*np.sum(yd*yp) - (np.sum(yd)*np.sum(yp))) / \
               (np.sqrt((n*(np.sum(np.square(yd))) - np.sum(yd)**2) * (n*(np.sum(np.square(yp))) - np.sum(yp)**2)))

    def plot(self):
        xd = self.t
        yd = self.p
        x = np.linspace(xd.min(), xd.max(), 100)
        popt = self.popt

        lpb, upb = self.exp_pb()
        lcb, ucb = self.exp_cb()

        plt.plot(x, ucb, c='k', alpha=0.35)
        plt.plot(x, lcb, c='k', alpha=0.35)
        plt.plot(x, self.exp_model(x, *popt), c='k', label='Exp. Fit')
        plt.plot(x, upb, c='r', alpha=0.5)
        plt.plot(x, lpb, c='r', alpha=0.5)
        plt.fill_between(x, y1=ucb, y2=lcb, color='k', alpha=0.20, label='95% Confidence Band')
        plt.fill_between(x, y1=upb, y2=lpb, color='r', alpha=0.20, label='95% Prediction Band')
        plt.scatter(xd, yd, alpha=0.10, label='APT Raw Data')
        plt.ylabel('Pressure (PSI)')
        plt.xlabel('time (hours)')
        plt.title('Kenda Air Permeation Test: {}'.format(self.name))
        plt.legend()
        plt.show()


class AirPermThermocouples:
    def __init__(self, **kwargs):
        if kwargs:
            for key, value in kwargs.items():
                setattr(self, key, value)


def main(input_file, output_file='output.xls'):
    log.info('CALLING read_xlsx(input_file={})'.format(input_file))
    apt = read_xlsx(input_file)
    # log.debug('{}'.format(apt.avg_temp_kelvin[0:5]))
    # log.debug('{}'.format(apt.avg_temp_celsius[0:5]))
    # log.debug('-'*50)
    # log.debug(apt.specs[0].name)
    # log.debug(apt.specs[0].t[20:25])
    # log.debug(apt.specs[0].normalized[20:25])
    log.debug('# of Specs: {}'.format(len(apt.specs)))
    log.debug('-'*50)
    # for spec in apt.specs:
    #     log.debug('Plotting Spec: {}'.format(spec.name))
    #     log.debug('Exp. Fit,   R: {}'.format(spec.exp_correlation()))
    #     spec.plot()
    #     spec.exp_kmpfit()
    #     os.system("pause")
    #     log.debug('-'*50)
    write_xlsx(apt, output_file)


def read_xlsx(file):
    START_ROW = 10

    log.debug('-'*50)
    log.debug('READING INPUT *.XLSX FILE')
    log.debug('-'*50)
    data = {}
    dates = []
    times = []
    wb = xlrd.open_workbook(file)
    log.debug("number of sheets: {}".format(wb.nsheets))
    for i in range(wb.nsheets):
        ws = wb.sheet_by_index(i)
        log.debug("worksheet {}: {}".format(i+1, ws.name))
        # log.debug('-'*50)
        # log.debug('READING COLUMN LABELS, BUILDING DATA DICTIONARY')
        # log.debug('-'*50)
        for j in range(ws.ncols):
            # log.debug("{}".format(ws.col_values(j, 0, 3)))
            if j == 0:
                dates = [datetime.datetime(*xlrd.xldate_as_tuple(val, wb.datemode)) for val in ws.col_values(j, START_ROW)]

            elif j == 1:
                times = [datetime.time((int(val * 24 * 3600)//3600),        # hours
                              (int(val * 24 * 3600) % 3600)//60,   # minutes
                              (int(val * 24 * 3600) % 60))
                         for val in ws.col_values(j, START_ROW)]
            else:
                if isinstance(ws.cell_value(1, j), (int, float)):
                    key = str(ws.cell_value(0, j)) + '_s' + str(int(ws.cell_value(1, j)))
                else:
                    key = str(ws.cell_value(0, j)) + '_' + str(ws.cell_value(1, j))
                key = key.replace('[', '').replace(']', '').replace('°', 'deg_').replace('.', '_').strip('_')
                data[key.lower()] = ws.col_values(j, START_ROW)
                pass

        if dates and times:
            data['datetime'] = [datetime.datetime.combine(date, times[i]) for i, date in enumerate(dates)]
            # log.debug('TOTAL SECONDS OF TIME 0, {}'.format(time.mktime(data['datetime'][0].timetuple())))
    # log.debug('-'*50)
    # log.debug('LOGGING DICTIONARY KEY VALUES')
    # log.debug('-'*50)
    # for key in dict.keys(data):
    #     log.debug(key)
        # if key == 'temp':
            # log.debug("{}".format([key for key in dict.keys(data['temp'])]))
    return AirPermTest(**data)


def write_xlsx(data, file):
    wb = xlwt.Workbook()
    ws1 = wb.add_sheet("Summary")
    # ws1.row(1).height_mismatch = True
    # ws1.row(1).height = 44*20

    init_row_heights = [15, 43, 27, 36, 18, 18]
    for i, h in enumerate(init_row_heights):
        ws1.row(i).height_mismatch = True
        ws1.row(i).height = h*20

    init_col_widths = [8, 8, 8, 17, 8]
    for i, w in enumerate(init_col_widths):
        ws1.col(i).width = w*367

    spec_sheets = list()
    for spec in data.specs:
        tmp = wb.add_sheet(spec.name)

        tmp.write(0, 1, label='Time (hrs)')
        tmp.write(0, 2, label='Nominal Pressure')
        tmp.write(0, 3, label='Normalized Pressure')
        tmp.write(0, 4, label='Averaged Temperature')
        for i, val in enumerate(spec.nom_p):
            tmp.write(i+1, 1, spec.t[i])
            tmp.write(i+1, 2, spec.nom_p[i])
            tmp.write(i+1, 3, spec.p[i])
            tmp.write(i+1, 4, spec.temps[i])
        spec_sheets.append(tmp)

    # summary label
    ws1.write_merge(1, 1, 1, 10,
                    label='AIR PERMEATION TESTING - SUMMARY',
                    style=xlwt.easyxf('font: height 360; align: horz center, vert center'))

    # tested specs label
    ws1.write_merge(2, 5, 1, 2,
                    label='TESTED SPECS',
                    style=xlwt.easyxf('font: bold 1, height 320; align: horz center, vert center, wrap 1'))

    # start date & time label
    ws1.write_merge(2, 2, 4, 5,
                    label='START',
                    style=xlwt.easyxf('font: bold 1; align: horz center, vert center;'))

    wb.save(file)


def init_logger():
    filelog = logging.getLogger('status')
    filelog.setLevel(logging.DEBUG)

    formatter = logging.Formatter('%(name)s - %(funcName)15s - %(lineno)d - %(levelname)s - %(message)s')

    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)

    fh = logging.FileHandler(filename='kpat_results.log', mode='a')
    fh.setLevel(logging.DEBUG)

    ch.setFormatter(formatter)
    fh.setFormatter(formatter)

    filelog.addHandler(fh)
    filelog.addHandler(ch)
    return filelog


if __name__ == "__main__":

    fmt = '%Y-%m-%d | %H:%M:%S'

    log = init_logger()
    log.info('==============================================================================')
    log.info('----------------          STARTING-UP THIS BIATCH         --------------------')
    log.info('-------------  BICYCLE GROUP AIR-PERM TEST, POST PROCESSING  -----------------')
    log.info('----------------           {}          --------------------'.format(
        str(datetime.datetime.now(timezone('US/Eastern')).strftime(fmt))
    ))
    log.info('==============================================================================')

    if len(sys.argv) >= 3:
        main(input_file=sys.argv[1], output_file=sys.argv[2])

    elif len(sys.argv) == 2:
        main(input_file=sys.argv[1])

    else:
        log.error("INCOMPETENCE ERROR: No input file, SOLUTION: I suggest to stop being stupid")
        log.info('PROGRAM TERMINATED WITH ERRORS - USER IS BEING A MORON')

    log.info('==============================================================================')
    log.info('----------------            {}         --------------------'.format(
        str(datetime.datetime.now(timezone('US/Eastern')).strftime(fmt))
    ))
    log.info('----------------                    DONE                  --------------------')
    log.info('-------------  BICYCLE GROUP AIR-PERM TEST, POST PROCESSING  -----------------')
    log.info('----------------                 PROGRAM END              --------------------')
    log.info('==============================================================================')