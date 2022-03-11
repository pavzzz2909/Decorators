from datetime import datetime
import pandas as pd

def insert_data(dict,arg1):
    df = pd.read_excel(arg1,sheet_name = 'logs')
    data = []
    data.append(list(dict.values()))
    new = pd.DataFrame(data, columns=dict.keys())
    res = pd.concat([df,new], ignore_index=True)
    res.to_excel(arg1, sheet_name = 'logs', index=False)

def decor_arg(arg1):
    def decorator(start_function):
        def do_function(*args, **kwargs):
            if args != ():
                result = start_function(*args, **kwargs)
                dict = {"Вызванная функция" : start_function.__name__,
                        "вызванный аргумент" : args[0],
                        "дата и время" : datetime.now().isoformat(),
                        "Результат" : result}
                insert_data(dict,arg1)
                return result
            else:
                result = start_function(*args, **kwargs)
                dict = {"Вызванная функция" : start_function.__name__,
                        "вызванный аргумент" : arg1,
                        "дата и время" : datetime.now().isoformat(),
                        "Результат" : result}
                insert_data(dict,arg1)
                return result
        return do_function
    return decorator
