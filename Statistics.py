"""
Excel-compatible statistics functions.

@author: Ben Hoffstein
@organization: Red Hook Software, LLC
@version: 1.0
@since: 2008-09-11
"""

import math

def average(data):
    """Returns the arithmetic mean."""
    n = len(data)
    if n == 0:
        raise Exception('One or more data points must exist.')

    return sum(data) / float(n)


def avgdev(data):
    """
    Returns the average of the absolute deviations of data points from their
    mean.
    """
    n = len(data)
    if n == 0:
        raise Exception('One or more data points must exist.')

    mu = average(data)
    tot = 0.0
    for x in data:
        tot += abs(x - mu)

    return (1 / float(n)) * tot


def correl(array1, array2):
    """
    Returns the correlation coefficient of the array1 and array2 cell ranges.
    """
    n1 = len(array1)
    n2 = len(array2)
    if n1 == 0 or n2 == 0:
        raise Exception('Arguments cannot be empty.')
    if n1 != n2:
        raise Exception('Arguments must have the same number of data points.')

    stdev1 = stdevp(array1)
    stdev2 = stdevp(array2)
    if stdev1 == 0.0 or stdev2 == 0.0:
        raise Exception('Standard deviation of an argument cannot be zero.')

    return covar(array1, array2) / (stdev1 * stdev2)


def covar(array1, array2):
    """
    Returns covariance, the average of the products of deviations for each data
    point pair.
    """
    n1 = len(array1)
    n2 = len(array2)
    if n1 == 0 or n2 == 0:
        raise Exception('Arguments cannot be empty.')
    if n1 != n2:
        raise Exception('Arguments must have the same number of data points.')

    return sumdev(array1, array2) / float(n1)


def devsq(data):
    """
    Returns the sum of squares of deviations of data points from their sample
    mean.
    """
    mu = average(data)
    tot = 0.0
    for x in data:
        tot += (x - mu) ** 2

    return tot


def forecast(x, known_y, known_x):
    """
    Calculates, or predicts, a future value by using existing values. The
    predicted value is a y-value for a given x-value. The known values are
    existing x-values and y-values, and the new value is predicted by using
    linear regression.
    """
    if isnumber(x) == False:
        raise Exception('X must be a numeric value.')

    varx = varp(known_x)
    if varx == 0:
        raise Exception('The variance of known_x cannot equal zero.')

    b = slope(known_y, known_x)
    a = intercept(known_y, known_x)

    return a + (b * x)


def intercept(known_y, known_x):
    """
    Calculates the point at which a line will intersect the y-axis by using
    existing x-values and y-values. The intercept point is based on a best-fit
    regression line plotted through the known x-values and known y-values. 
    """
    muy = average(known_y)
    mux = average(known_x)
    b = slope(known_y, known_x)

    return muy - (b * mux)

    
def isnumber(x):
    """Returns True if x is a number or False if not."""
    return hasattr(x, '__int__')


def pearson(array1, array2):
    """
    Returns the Pearson product moment correlation coefficient, r, a
    dimensionless index that ranges from -1.0 to 1.0 inclusive and reflects the
    extent of a linear relationship between two data sets.
    """
    return correl(array1, array2)


def rsq(known_y, known_x):
    """
    Returns the square of the Pearson product moment correlation coefficient
    through data points in known_y's and known_x's.
    """
    return correl(known_y, known_x) ** 2


def slope(known_y, known_x):
    """
    Returns the slope of the linear regression line through data points in
    known_y's and known_x's.  The slope is the vertical distance divided by the
    horizontal distance between any two points on the line, which is the rate
    of change along the regression line.
    """
    return sumdev(known_x, known_y) / devsq(known_x)


def sumdev(array1, array2):
    """
    Returns the sum of the deviations of data points from their sample mean.
    """
    n1 = len(array1)
    n2 = len(array2)
    if n1 == 0 or n2 == 0:
        raise Exception('Arguments cannot be empty.')
    if n1 != n2:
        raise Exception('Arguments must have the same number of data points.')

    mu1 = average(array1)
    mu2 = average(array2)

    tot = 0.0
    for (x, y) in zip(array1, array2):
        tot += ((x - mu1) * (y - mu2))

    return tot


def var(data):
    """
    Returns the variance, assuming the data represents a sample of the
    population.
    """
    return variance(data, True)


def varp(data):
    """
    Returns the variance, assuming the data represents the entire
    population.
    """
    return variance(data, False)


def variance(data, is_sample=True):
    """Returns the variance."""
    n = len(data)
    if n == 0:
        raise Exception('One or more data points must exist.')

    mu = average(data)
    ds = devsq(data)

    if is_sample:
        return ds / float(n - 1)
    return ds / float(n)


def stdev(data):
    """
    Returns the standard deviation, assuming the data represents a sample
    of the population.
    """
    return standard_deviation(data, True)


def stdevp(data):
    """
    Returns the standard deviation, assuming the data represents the entire
    population.
    """
    return standard_deviation(data, False)


def standard_deviation(data, is_sample=True):
    """Returns the standard deviation."""
    if is_sample:
        return math.sqrt(var(data))
    return math.sqrt(varp(data))

