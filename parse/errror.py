
class ErrorNotValidExcel(Exception):

    def __init__(self):
        pass


def with_error_stack(e):
    return {
        'error': e,
        'file': e.__traceback__.tb_frame.f_globals['__file__'],
        'line': e.__traceback__.tb_lineno,
    }
