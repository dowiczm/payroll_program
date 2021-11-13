import pandas as pd

def to_float(df):
    float_columns = ['Quantity', 'Charge', 'Discount', 'Net', 'COGS', 'Profit', 'Seller', 'Provider', 'Manager', 'Gratuity']

    for column in df.columns:    #cleaning services
        if df[column].dtype == 'object' and column in float_columns:
            df[column] = [float(x.replace(',', '')) for x in df[column]]
    return

