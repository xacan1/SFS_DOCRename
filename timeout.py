from functools import wraps
import errno
import os
from threading import Timer


def timeout(seconds=10, error_message=os.strerror(errno.ETIME)):
    def decorator(func):
        @wraps(func)
        def _handle_timeout(*args, **kwargs):
            def _raise_timeout():
                raise TimeoutError
            timer = Timer(seconds, _raise_timeout)
            timer.start()
            try:
                result = func(*args, **kwargs)
            finally:
                timer.cancel()
            return result
        return _handle_timeout
    return decorator
