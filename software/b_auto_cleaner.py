from datetime import datetime
import pandas as pd

def auto_cleaner(df):
    df_list = []
    
    indexs = [(n,x) for n,x in enumerate(df['Commissions and Gratuity']) if x == 'Date']
    df_length = len(df)
    index_length = len(indexs)
    
    for i in range(0, index_length-1):
        first = indexs[i][0] - 1
        second = indexs[i+1][0]
        
        piece = df[first:second].reset_index(drop=True)
        piece.columns = ['Date', 'Patient', 'Procedure/Product',
                    'Quantity', 'Charge', 'Discount', 'Net',
                    'COGS', 'Profit', 'Seller', 'Provider',
                    'Manager', 'Gratuity']
        
        
        name = piece['Date'][0].split(":")[1]
        clean_name = ' '.join(name.split(' ')[1:])
        
        piece['Employee'] = clean_name
        
        piece.fillna(value='None', inplace=True)
        
        piece = piece[piece['Date'].str.contains('21')]
        
        
        df_list.append(piece)
    
    
    
    
    last = df[indexs[-1][0]-1:df_length-1].reset_index(drop=True)
 
    last.columns = ['Date', 'Patient', 'Procedure/Product',
                    'Quantity', 'Charge', 'Discount', 'Net',
                    'COGS', 'Profit', 'Seller', 'Provider',
                    'Manager', 'Gratuity']
        
        
    name = last['Date'][0].split(":")[1]
    clean_name = ' '.join(name.split(' ')[1:])
        
    last['Employee'] = clean_name
        
    last.fillna(value='None', inplace=True)
        
    last = last[last['Date'].str.contains('21')]
    
    df_list.append(last)
    
        
        
        
    return pd.concat(df_list)
