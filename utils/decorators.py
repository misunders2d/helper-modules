from ctk_gui.ctk_windows import PopupError
import sys

def error_checker(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            error_str = f'function {func.__name__} returned an error:\n{e}'
            PopupError(message=error_str)
            sys.exit(1)
    return wrapper