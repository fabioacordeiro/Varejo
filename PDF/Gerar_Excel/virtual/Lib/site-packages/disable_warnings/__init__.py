def _warnxXXXXXXXXX(*args, **kwargs):
    return ''
import warnings
warnings.warn = _warnxXXXXXXXXX
warnings.warn_explicit = _warnxXXXXXXXXX
warnings.showwarning = _warnxXXXXXXXXX