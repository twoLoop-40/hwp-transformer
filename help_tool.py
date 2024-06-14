import os


def get_home_directory_name() -> str:
    home_directory = os.path.expanduser("~")
    return home_directory


def identity(num: int) -> int:
    return num


def half(num: int) -> int:
    return int(num / 2)
