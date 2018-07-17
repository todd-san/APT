import time
import numpy as np

from scipy import optimize

CELSIUS_2_KELVIN = 273.15
IDEAL_APT_ROOM = 22.22222


class AptSpec:
    def __init__(self, name, trim=None, **kwargs):
        """
        :param name: parsed name of the measured spec
        :param trim: <Trim> object that has an attribute "start" and "end" for slicing signals
        :param kwargs: <dict> that should contain keys:
            psi: <list> of <str> objects that should be <float> convertible
            baro: <list> of <str> objects that should be <float> convertible
            datetime: <list> of <str> objects that should be <datetime> objects
            t#: <list> of <dict> objects where:
                    key: thermocouple name
                    value: <list> of <str> objects that should be <float> convertible
        """
        # predefined (expected class attributes)
        self.name = name
        self.psi = list()
        self.baro = list()
        self.datetime = list()
        self.temps = list()

        # populating instantiated variables,
        # handles a flexible amount of thermocouples
        for key, value in kwargs.items():
            if not trim:
                if len(key) == 2 and key[0] == 't':
                    self.temps.append({key: value})
                else:
                    setattr(self, key, value)
            else:
                if len(key) == 2 and key[0] == 't':
                    self.temps.append({key: value[trim.start:trim.end]})
                else:
                    setattr(self, key, value[trim.start: trim.end])

    @property
    def avg_temp(self):
        """
        Averages an array of thermocouple readings and converts them to Kelvin from Celsius
        :return: temp(K)
        """
        values = list()
        temps = list()
        for item in self.temps:
            for key in dict.keys(item):
                values = [float(value)+CELSIUS_2_KELVIN for value in item[key]]
            temps.append(values)
        return list(np.mean(np.array(temps), axis=0))

    @property
    def time(self):
        """
        Converts datetime object to a floating point number of hours
        :return: time (hours)
        """
        t0 = time.mktime(self.datetime[0].timetuple())
        t = [time.mktime(tN.timetuple()) - t0 for tN in self.datetime]
        return np.array([val/3600 for val in t])

    @property
    def pressure(self):
        p = list()
        for i, val in enumerate(self.baro):
            p.append((val*(IDEAL_APT_ROOM+CELSIUS_2_KELVIN))/(self.avg_temp[i]))
        return np.array(p)

    @property
    def curve_fit(self):
        try:
            popt, pcov = optimize.curve_fit(self.exp_model, self.time, self.pressure)
            return {
                'popt': popt,
                'pcov': pcov}
        except RuntimeError as re:
            return {
                'popt': np.array([0, 0, 0]),
                'pcov': np.array([[0, 0, 0],
                                  [0, 0, 0],
                                  [0, 0, 0]])
            }

    @staticmethod
    def exp_model(t, a, b, c):
        return a * np.exp(-b * t) + c

    def exp_confidence_bands(self):
        return self.name

    def exp_prediction_bands(self):
        return self.name
