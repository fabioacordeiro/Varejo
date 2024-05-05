from disable_warnings import *
from typing import Any
from collections import defaultdict
try:
    from collections import Iterable
except Exception:
    from collections.abc import Iterable
from functools import reduce
import numpy as np
import pandas as pd
import math


def checkiter(x):
    try:
        _ = iter(x)
        return True
    except TypeError:
        return False


def float_check_nan(num):
    if float("-inf") < float(num) < float("inf"):
        return False
    else:
        return True


def is_nan(
    x: Any,
    emptyiters: bool = False,
    nastrings: bool = False,
    emptystrings: bool = False,
    emptybytes: bool = False,
) -> bool:
    # useful when you read a csv file and the missing data is not converted correctly to nan
    nastringlist = [
        "<NA>",
        "<NAN>",
        "<nan>",
        "np.nan",
        "NoneType",
        "None",
        "-1.#IND",
        "1.#QNAN",
        "1.#IND",
        "-1.#QNAN",
        "#N/A N/A",
        "#N/A",
        "N/A",
        "n/a",
        "NA",
        "#NA",
        "NULL",
        "null",
        "NaN",
        "-NaN",
        "nan",
        "-nan",
    ]
    if isinstance(x, type(None)):
        return True
    try:
        if np.isnan(x):
            return True
    except Exception:
        pass
    try:
        if pd.isna(x):
            return True
    except Exception:
        pass
    try:
        if pd.isnull(x):
            return True
    except Exception:
        pass
    try:
        if math.isnan(x):
            return True
    except Exception:
        pass
    try:
        if x != x:
            return True
    except Exception:
        pass
    try:
        if not isinstance(x, str):
            if float_check_nan(x) is True:
                return True
    except Exception:
        pass
    if emptystrings is True:
        if isinstance(x, str):

            try:
                if x == "":
                    return True
            except Exception:
                pass
    if emptybytes is True:
        if isinstance(x, bytes):
            try:
                if x == b"":
                    return True
            except Exception:
                pass
    if nastrings is True:
        if isinstance(x, str):

            try:
                if x in nastringlist:
                    return True
            except Exception:
                pass
    if emptyiters is True:
        if isinstance(x, (str, bytes)):
            pass
        else:
            if isinstance(x, Iterable) or checkiter(x):
                try:
                    if not np.any(x):
                        return True
                except Exception:
                    pass
                if isinstance(x, (pd.Series, pd.DataFrame)):
                    try:
                        if x.empty:
                            return True
                    except Exception:
                        pass
                try:
                    if not any(x):
                        return True
                except Exception:
                    pass

                try:
                    if not x:
                        return True
                except Exception:
                    pass
                try:
                    if len(x) == 0:
                        return True
                except Exception:
                    pass
    return False


def groupBy(key, seq):
    # https://stackoverflow.com/a/60282640/15096247
    return reduce(
        lambda grp, val: grp[key(val)].append(val) or grp, seq, defaultdict(list)
    )


def sort_nan_non_nan(
    seq,
    emptyiters: bool = False,
    nastrings: bool = False,
    emptystrings: bool = False,
    emptybytes: bool = False,
):
    return groupBy(
        key=lambda x: True
        if is_nan(
            x[1],
            emptyiters=emptyiters,
            nastrings=nastrings,
            emptystrings=emptystrings,
            emptybytes=emptybytes,
        )
        else False,
        seq=(enumerate(seq)),
    )


