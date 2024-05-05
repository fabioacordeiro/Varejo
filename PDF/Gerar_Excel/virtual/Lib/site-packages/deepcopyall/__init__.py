from copy import deepcopy as dc_
import dill


def deepcopy(o):
    try:
        cpy = dc_(o)
        if cpy is not o:
            return cpy
    except Exception:
        pass
    try:
        cpy = dill.dumps(o)
        return dill.loads(cpy)
    except Exception:
        return None
