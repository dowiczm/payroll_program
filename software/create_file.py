import pandas as pd
import warnings
warnings.filterwarnings("ignore")
from datetime import datetime

import os


def create_file(services, products, medical_professionals, massage_therapists, ests, management, front_desk, all_employees, destination):
    date = datetime.now().strftime('%m-%d-%y')
    
    
    writer = pd.ExcelWriter(f'{destination}/Master Payroll {date}.xlsx', engine='xlsxwriter')
    
    
    #Cover Sheet
    name_list = []
    for group in all_employees:
        for name in group:
            name_list.append(name)
        name_list.append('-')
    
    name_df = pd.DataFrame(name_list, columns=['Name'])
    
    map_dict = {
    'Ae' : medical_professionals,
    'LMT' : massage_therapists,
    'Med' : ests,
    'FNT' : front_desk,
    'Mgt' : management
    }
    
    employee_type = []
    for group in map_dict:
        people = map_dict[group]
        for name in people:
            employee_type.append(group)
        employee_type.append('-')
    
    name_df['Type'] = employee_type
    name_df.to_excel(writer, sheet_name='Cover Sheet', index=False)
    

    #Individual Names
    for name in medical_professionals:
        df = services.loc[services['Employee'] == name]
        df.to_excel(writer, sheet_name=name, index=False)


    for name in massage_therapists:
        df = services.loc[services['Employee'] == name]
        new_quantities = []

        for procedure  in df['Procedure/Product']:   #need to get a list and map them instead
            if '80' in procedure:
                new_quantities.append(1.5)
            elif '30' in procedure:
                new_quantities.append(0.5)
            elif '15' in procedure:
                new_quantities.append(0.25)
            elif 'ADD ON' in procedure:
                new_quantities.append(0.5)
            else:
                new_quantities.append(1)

        df['Quantity'] = new_quantities

        df.to_excel(writer, sheet_name=name, index=False)


    for name in ests:
        df = services.loc[services['Employee'] == name]
        df.to_excel(writer, sheet_name=name, index=False)
    
    for name in management:
        df = services.loc[services['Employee'] == name]
        df.to_excel(writer, sheet_name=name, index=False)

    for name in front_desk:
        df = products.loc[products['Employee'] == name]
        df.to_excel(writer, sheet_name=name, index=False)
    
    

    
    
        
    
        
    writer.save()
    
    
    
    
    