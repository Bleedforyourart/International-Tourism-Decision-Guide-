import pandas as pd
import xlsxwriter
import tkinter as tk
from tkinter import filedialog
import os

application_window = tk.Tk()

#use tkinter file dialog to choose file
filename=filedialog.askopenfilename(parent=application_window, initialdir=os.getcwd(), filetypes=[("Excel files", ".xlsx")], title="Please select the International Tourist Arrivals Worldwide file: ")

#read dataframe for rows on overview sheet
df1a = pd.read_excel(filename, 'Overview', header=8, usecols=str("B:C"), skipfooter=1)
df1a=df1a.drop(labels=[4,5,6,7,8,9,10], axis=0)

#read dataframe for description on overview sheet
df1b = pd.read_excel(filename, 'Overview', header=6, usecols=str("E"))
df1b=df1b.drop(labels=[1], axis=0)

#read dataframe for description info on overview sheet
df1c = pd.read_excel(filename, 'Overview', header=7, usecols=str("E"))

#read dataframe for data sheet
df2 = pd.read_excel(filename, 'Data', header=4)

#write statistics variables for each regions mean, min and max 
americas_mean=df2['Americas'].mean()
americas_min=df2['Americas'].min()
americas_max=df2['Americas'].max()
asiapacific_mean=df2['Asia Pacific'].mean()
asiapacific_min=df2['Asia Pacific'].min()
asiapacific_max=df2['Asia Pacific'].max()
middleeast_mean=df2['Middle East'].mean()
middleeast_min=df2['Middle East'].min()
middleeast_max=df2['Middle East'].max()
africa_mean=df2['Africa'].mean()
africa_min=df2['Africa'].min()
africa_max=df2['Africa'].max()
europe_mean=df2['Europe'].mean()
europe_min=df2['Europe'].min()
europe_max=df2['Europe'].max()

#write the xlsx file
writer=pd.ExcelWriter('international_tourism_data.xlsx',engine='xlsxwriter')

#write the dataframes to each new sheet
df1a.to_excel(writer, sheet_name='Overview', header=False,index=False)
df1b.to_excel(writer,sheet_name='Overview',header=False,index=False, startrow=int(8))
df1c.to_excel(writer,sheet_name='Overview',header=False,index=False, startrow=int(8), startcol=int(1))
df2.to_excel(writer,sheet_name='Data', index=False)

#create workbook and worksheets
workbook = writer.book
worksheet1 = writer.sheets['Overview']
worksheet2 = writer.sheets['Data']
worksheet3 = workbook.add_worksheet('Statistics')

#write formatting variables
format = workbook.add_format()
format.set_align('left')
format.set_bold()

#format overview worksheet
worksheet1.set_column('A:A', 30, format)
worksheet1.set_column('B:B', 100)
worksheet1.set_default_row(20)

#format data worksheet
worksheet2.set_column('A:F', 30)
worksheet2.set_default_row(20)

#format statistics worksheet
worksheet3.set_column('A:F', 30)
worksheet3.set_default_row(20)

#add headings and format for statistics sheet
worksheet3.write("A1", 'Statistics by Region', format)
worksheet3.write("B1", 'Americas', format)
worksheet3.write("C1", 'Asia Pacific', format)
worksheet3.write("D1", 'Middle East', format)
worksheet3.write("E1", 'Africa', format)
worksheet3.write("F1", 'Europe', format)
worksheet3.write("A2", 'Mean', format)
worksheet3.write("A3", 'Min', format)
worksheet3.write("A4", 'Max', format)

#add data for statistics sheet
worksheet3.write("B2", americas_mean)
worksheet3.write("B3", americas_min)
worksheet3.write("B4", americas_max)
worksheet3.write("C2", asiapacific_mean)
worksheet3.write("C3", asiapacific_min)
worksheet3.write("C4", asiapacific_max)
worksheet3.write("D2", middleeast_mean)
worksheet3.write("D3", middleeast_min)
worksheet3.write("D4", middleeast_max)
worksheet3.write("E2", africa_mean)
worksheet3.write("E3", africa_min)
worksheet3.write("E4", africa_max)
worksheet3.write("F2", europe_mean)
worksheet3.write("F3", europe_min)
worksheet3.write("F4", europe_max)

writer.save()
workbook.close()
