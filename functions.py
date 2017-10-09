import statistics
from functools import wraps

def get_all_items(items):
    final = list()
    if any([isinstance(items, typo) for typo in (list, tuple)]):
        for item in items:
            final.extend(get_all_items(item))
    else:
        final.append(items)
    return final

def fx(function):
    @wraps(function)
    def wrapper(*args, **kwargs):
        args = get_all_items(args)
        try:
            data = function(*args, **kwargs)
        except (TypeError, ValueError):
            def check_int(value):
                try:
                    int(value)
                except (TypeError, ValueError):
                    return False
                else:
                    return True
            args = list(filter(check_int, args))
            data = function(*args, **kwargs)
        return data
    return wrapper

class Functions:
    @staticmethod
    @fx
    def average(*args):
        return statistics.mean(args)

    @staticmethod
    @fx
    def count(*args):
        return len(list(filter(lambda x: x not in (None, ""), args)))

    @staticmethod
    @fx
    def max(*args):
        return max(args)

    @staticmethod
    @fx
    def min(*args):
        return min(args)

    @staticmethod
    @fx
    def sum(*args):
        print(args)
        return sum(args)
