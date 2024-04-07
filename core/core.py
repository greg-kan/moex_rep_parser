import os


def ifnull(var, val):
    if var is None:
        return val
    return var


def tofloat(var) -> float | None:
    if var is None:
        return None

    try:
        return float(var)
    except Exception as e:
        return None


def get_file_name(file_name):
    f_name, _ = os.path.splitext(file_name)
    return f_name


def get_file_extension(file_name):
    _, f_extension = os.path.splitext(file_name)
    return f_extension

