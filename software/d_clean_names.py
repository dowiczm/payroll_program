import pandas as pd

def clean_employee_names(df):   
    new_names = []
    employees = list(df['Employee'])
    for name in employees:
        base_name = ''
        for x in name.split(' '):
            if len(x) > 0:
                base_name += x + ' '
        cleaned_name = base_name[:-1]
        new_names.append(cleaned_name)
    df['Employee'] = new_names
    
    
