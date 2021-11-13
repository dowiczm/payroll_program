import pandas as pd
from datetime import datetime
import openpyxl

import warnings
warnings.filterwarnings("ignore")

from b_auto_cleaner import auto_cleaner
from c_to_float import to_float
from d_clean_names import clean_employee_names
from create_file import create_file
from master import master_payroll

import tkinter as tk
from tkinter import filedialog
from tkinter.filedialog import askopenfile

from PIL import Image, ImageTk

import os

#User Path
user_path = os.getcwd()
folder_path = '\Desktop\Payroll\software\jpg\\'


#Creating Canvas
HEIGHT = 400
WIDTH = 600

title_font = ('times', 12, 'bold')
color = '#009999'

root = tk.Tk()
root.title('New Beauty & Wellness Payroll Software')

canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH)
canvas.pack()

img = ImageTk.PhotoImage(Image.open(f'{user_path}{folder_path}landscape.png'))
background_label = tk.Label(root, image=img)
background_label.place(x=0,y=0,relwidth=1, relheight=1)



#Service Frame
service_str = tk.StringVar()
service_str.set("Upload Service File -->")


service_frame = tk.Frame(root, bg=color, bd=5)
service_frame.place(relx=0.5, rely=0.1, relwidth=0.5, relheight=0.1, anchor='n')

service_label = tk.Label(service_frame, bg=color, font=title_font, textvariable=service_str)
service_label.place(relwidth=0.7, relheight=0.8)

service_button = tk.Button(service_frame, text='Service',command=lambda:upload_service_file())
service_button.place(relx=0.7,relheight=1, relwidth=0.3)

service_path = []


def upload_service_file():
    file = filedialog.askopenfilename()
    
    if (file):
        service_str.set(f"File Uploaded: {file.split("\\")[-1]}")
        file=open(file,'r')     
        service_path.append(f'{file.name}')

        
        
#Product Frame
product_str = tk.StringVar()
product_str.set('Upload Product File -->')

product_frame = tk.Frame(root, bg=color, bd=5)
product_frame.place(relx=0.5, rely=0.25, relwidth=0.5, relheight=0.1, anchor='n')

product_label = tk.Label(product_frame, bg=color, font=title_font, textvariable=product_str)
product_label.place(relwidth=0.7, relheight=0.8)

product_button = tk.Button(product_frame, text='Product', command=lambda:upload_product_file())
product_button.place(relx=0.7,relheight=1, relwidth=0.3)

product_path = []

def upload_product_file():
    file = filedialog.askopenfilename()
    
    if (file):
        product_str.set(f"File Uploaded: {file.split('\\')[-1]}")
        file=open(file, 'r')
        product_path.append(f'{file.name}')




#Info Frame
info_str = tk.StringVar()
info_str.set('Upload Info File -->')

info_frame = tk.Frame(root, bg=color, bd=5)
info_frame.place(relx=0.5, rely=0.4, relwidth=0.5, relheight=0.1, anchor='n')

info_label = tk.Label(info_frame, bg=color, font=title_font, textvariable=info_str)
info_label.place(relwidth=0.7, relheight=0.8)

info_button = tk.Button(info_frame, text='Info', command=lambda:upload_info_file())
info_button.place(relx=0.7,relheight=1, relwidth=0.3)

info_path = []

def upload_info_file():
    file = filedialog.askopenfilename()
    
    if (file):
        info_str.set(f"File Uploaded: {file.split('\\')[-1]}")
        file=open(file, 'r')
        info_path.append(f'{file.name}')




#Set Destination Frame
dest_str = tk.StringVar()
dest_str.set('Set Destination -->')

dest_frame = tk.Frame(root, bg=color, bd=5)
dest_frame.place(relx=0.5, rely=0.55, relwidth=0.5, relheight=0.1, anchor='n')

dest_label = tk.Label(dest_frame, bg=color, font=title_font, textvariable=dest_str)
dest_label.place(relwidth=0.7, relheight=0.8)

dest_button = tk.Button(dest_frame, text='Folder', command=lambda:set_destination())
dest_button.place(relx=0.7, relheight=1, relwidth=0.3)

destination_path = []

def set_destination():
    file = filedialog.askdirectory()
    
    if (file):
        dest_str.set(f"Destination: {file.split('\\')[-1]}")    
        destination_path.append(f'{file}')


        
#Generate Report Button
date = datetime.now().strftime('%m-%d-%y')

generate_button = tk.Button(root, text='Generate Report', command=lambda:master_payroll(service_path, product_path, info_path, destination_path, date))
generate_button.place(relx=0.5, rely=0.8, relwidth=0.25, relheight=0.1, anchor='n')


root.mainloop()