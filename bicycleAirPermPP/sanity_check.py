import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit


def fitFunc(t, a, b, c):
    return a * np.exp(-b * t) + c


t = np.linspace(0, 4, 50)
temp = fitFunc(t, 2.5, 1.3, 0.5)
noisy = temp + 0.25*np.random.normal(size=len(temp))

fitParams, fitCovariance = curve_fit(fitFunc, t, noisy)

print(fitParams)
print(fitCovariance)

plt.figure()
plt.ylabel('Pressure (PSI)', fontsize=16)
plt.xlabel('time (s)', fontsize=16)
plt.xlim(0, 4.1)
# plot the data as red circles with vertical errorbars
plt.errorbar(t, noisy, fmt='ro', yerr=0.2)
# now plot the best fit curve and also +- 1 sigma curves
# (the square root of the diagonal covariance matrix
# element is the uncertianty on the fit parameter.)

sigma = [fitCovariance[0, 0], fitCovariance[1, 1], fitCovariance[2, 2]]
plt.plot(t, fitFunc(t, fitParams[0], fitParams[1], fitParams[2]),
         t, fitFunc(t, fitParams[0] + sigma[0], fitParams[1] - sigma[1], fitParams[2] + sigma[2]),
         t, fitFunc(t, fitParams[0] - sigma[0], fitParams[1] + sigma[1], fitParams[2] - sigma[2]))

plt.show()