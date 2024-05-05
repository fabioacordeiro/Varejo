# Checks for all kinds of nan/None values without raising Exceptions all the time


```python
from check_if_nan import is_nan,sort_nan_non_nan
import numpy as np
import pandas as pd
import math
a = None
b = pd.NA
c = np.nan
d = math.nan
e = float("nan")
f = []
g = np.array([])
h = dict()
i = tuple()
j = set()
k = ""
l = "NaN"
m = b""
n = bytearray()


print("a", is_nan(a))
print("b", is_nan(b))
print("c", is_nan(c))
print("d", is_nan(d))
print("e", is_nan(e))
print("f", is_nan(f))
print("g", is_nan(g))
print("h", is_nan(h))
print("i", is_nan(i))
print("j", is_nan(j))
print("k", is_nan(k))
print("l", is_nan(l))
print("m", is_nan(m))
print("n", is_nan(n))

print("f", is_nan(f, emptyiters=True))
print("g", is_nan(g, emptyiters=True))
print("h", is_nan(h, emptyiters=True))
print("i", is_nan(i, emptyiters=True))
print("j", is_nan(j, emptyiters=True))
print("k", is_nan(k, emptystrings=True))
print("l", is_nan(l, nastrings=True))
print("m", is_nan(m, emptybytes=True))
print("n", is_nan(n, emptyiters=True))


a True
b True
c True
d True
e True
f False
g False
h False
i False
j False
k False
l False
m False
n False


f True
g True
h True
i True
j True
k True
l True
m True
n True


sor = sort_nan_non_nan(
    seq=[a, b, c, d, e, f, g, h, i, j, k, l, m, n],
    emptyiters=False,
    nastrings=False,
    emptystrings=False,
    emptybytes=False,
)
print(sor)
# defaultdict(<class 'list'>, {True: [(0, None), (1, <NA>), (2, nan),
# (3, nan), (4, nan)], False: [(5, []), (6, array([], dtype=float64)),
# (7, {}), (8, ()), (9, set()), (10, ''), (11, 'NaN'), (12, b''),
# (13, bytearray(b''))]})

sor = sort_nan_non_nan(
    seq=[a, b, c, d, e, f, g, h, i, j, k, l, m, n],
    emptyiters=True,
    nastrings=False,
    emptystrings=False,
    emptybytes=False,
)
print(sor)
# defaultdict(<class 'list'>, {True: [(0, None), (1, <NA>), (2, nan),
# (3, nan), (4, nan), (5, []), (6, array([], dtype=float64)),
# (7, {}), (8, ()), (9, set()), (13, bytearray(b''))],
# False: [(10, ''), (11, 'NaN'), (12, b'')]})


sor = sort_nan_non_nan(
    seq=[a, b, c, d, e, f, g, h, i, j, k, l, m, n],
    emptyiters=True,
    nastrings=False,
    emptystrings=True,
    emptybytes=True,
)
print(sor)
# defaultdict(<class 'list'>, {True: [(0, None), (1, <NA>), (2, nan), (3, nan),
# (4, nan), (5, []), (6, array([], dtype=float64)), (7, {}), (8, ()),
# (9, set()), (10, ''), (12, b''), (13, bytearray(b''))], False: [(11, 'NaN')]})

sor = sort_nan_non_nan(
    seq=[a, b, c, d, e, f, g, h, i, j, k, l, m, n],
    emptyiters=True,
    nastrings=True,
    emptystrings=True,
    emptybytes=True,
)
print(sor)
# defaultdict(<class 'list'>, {True: [(0, None), (1, <NA>), (2, nan),
# (3, nan), (4, nan), (5, []), (6, array([], dtype=float64)), (7, {}),
# (8, ()), (9, set()), (10, ''), (11, 'NaN'), (12, b''), (13, bytearray(b''))]})
```


