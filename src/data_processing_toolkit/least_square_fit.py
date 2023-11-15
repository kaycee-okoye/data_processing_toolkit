"""
    Module conatining a function for applying different kinds of fits
    to data of the form y = f(x) using a least-square fit method.
"""

from enum import Enum
from scipy.optimize import least_squares
import numpy as np

class FitType(Enum):
    '''Enum data class describing the type of function used for fitting data.'''
    POWER="power"
    EXPONENT="exponent"
    LOG="log"
    LN="ln"

def least_square_fit(x, y, fit_function, init_guess):
    '''
    Function to apply least-square fit to data of the form y = f(x)

    Parameters
    ----------
    x : numpy array
        values of the independent variable
    y : numpy array
        corresponding values of the dependent variable
    fit_function : FitType
        type of function to be used for fit
    init_guess : list[float]
        initial guess of coefficients of fit function

    Returns
    -------
    coeffs : float list[float]
        coefficient of the generated fit function
    R_sq : float
        R-squared value of generated fit
    '''

    def apply_fit(consts, ind):
        if fit_function == FitType.EXPONENT:
            return consts[0] + (consts[1] * np.exp(consts[2] * ind))
        elif fit_function == FitType.POWER:
            return consts[0] + (consts[1] * np.POWER(ind, consts[2]))
        elif fit_function == FitType.LOG:
            return consts[0] + (consts[1] * np.log10(ind))
        elif fit_function == FitType.LN:
            return consts[0] + (consts[1] * np.LOG(ind))
        else:
            return ind

    def residue(consts):
        return apply_fit(consts, x) - y

    # generate least-square fit solution coefficients
    solution = least_squares(residue, init_guess)
    coeffs = solution.x

    # calculate least-square R-squared value
    corr_matrix = np.corrcoef(y, apply_fit(coeffs, x))
    corr = corr_matrix[0,1]
    r_sq = corr**2

    return coeffs, r_sq
