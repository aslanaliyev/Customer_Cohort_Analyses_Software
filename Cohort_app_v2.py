# Created by Aslan Aliyev 

import sys
import tkinter as tk
import os
import xlrd
import pandas as pd
import numpy as np
from pathlib import Path
from glob import glob
from functools import reduce
from functools import partial, reduce
from tkinter import filedialog
from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
# from dfply import *
import datetime
from pandas import Series, ExcelWriter
import math

desired_width = 300

pd.set_option('display.width', desired_width)

np.set_printoptions(linewidth=desired_width)

pd.set_option('display.max_columns', 15)

#############################################################
#############################################################
###################### Tkinter app ##########################


window = Tk()
window.title("Cohort App (beta ver:1.01)")
window.geometry("500x300")

lbl_0 = Label(window, text="")
lbl_0.grid(column=0, row=0)
lbl_1 = Label(window, text="             Bitte Spalte als Zahl angeben!!! (z.B. A=1, B=2, C=3)",
              font='Helvetica 9 bold')
lbl_1.grid(column=0, row=1)
lbl_2 = Label(window, text="\t Spalte für Kundennummer")
lbl_2.grid(column=0, row=2)
lbl_4 = Label(window, text="\t Spalte für Umsatz")
lbl_4.grid(column=0, row=3)
lbl_6 = Label(window, text="\t Spalte für Jahr")
lbl_6.grid(column=0, row=4)
lbl_7 = Label(window, text="\t Spalte für Umsatzart (falls gibt!)")
lbl_7.grid(column=0, row=5)
lbl_8 = Label(window, text="Developed by: Aslan Aliyev (a.aslan1991@gmail.com)", font='Helvetica 8')
lbl_8.grid(sticky="W", column=0, row=15, padx=0, pady=65)

txt1 = Entry(window, width=10)
txt1.grid(column=1, row=2, padx=5, pady=5)
txt2 = Entry(window, width=10)
txt2.grid(column=1, row=3, padx=5, pady=5)
txt3 = Entry(window, width=10)
txt3.grid(column=1, row=4, padx=5, pady=5)
txt4 = Entry(window, width=10)
txt4.grid(column=1, row=5, padx=5, pady=5)


###################################################
###################################################

def clicked():
    global Kundennummer
    global Umsatz
    global Jahr

    Kundennummer = int(txt1.get())
    Kundennummer = Kundennummer - 1
    Umsatz = int(txt2.get())
    Umsatz = Umsatz - 1
    Jahr = int(txt3.get())
    Jahr = Jahr - 1
    # Umsatzart = int(txt4.get())
    # if len(txt4.get()) != 0:
    #     Umsatzart = int(txt4.get())
    #     Umsatzart = Umsatzart - 1

    data = pd.read_excel(folder_path)

    Kundennummer_new = data.columns[Kundennummer]
    Umsatz_new = data.columns[Umsatz]
    Jahr_new = data.columns[Jahr]


    # Umsatzart_new = data.columns[Umsatzart]

    if (len(str(data.loc[1, Jahr_new])) > 4 and len(str(data.loc[50, Jahr_new])) > 4 and len(
            str(data.loc[25, Jahr_new])) > 4):
        data[Jahr_new] = pd.DatetimeIndex(data[Jahr_new]).year.copy()
    # Umsatzart_list = data[Umsatzart_new].unique().tolist()
    # Umsatzart_list_length = len(data[Umsatzart_new].unique().tolist())
    data[Umsatz_new] = pd.to_numeric(data[Umsatz_new], errors='coerce')
    data.sort_values(by=[Umsatz_new], inplace=True)
    data = data[(data[Umsatz_new].notnull()) & (data[Umsatz_new] != 0)].copy()
    data_umsatzart1 = data.copy()
    data.set_index(Kundennummer_new, inplace=True)

    data["cohort_group"] = data.groupby(level=0)[Jahr_new].min()
    data1 = data.copy()

    # data_dpl_final = data_dpl >> group_by(Kundennummer_new) >> mutate(first = min(Jahr_new)) >>  \
    #                 group_by(first, Jahr_new) >> summarize( Unique = Kundennummer_new.n_unique()) >>         \
    #                 spread(Jahr_new, Kunden) >> n_distinct

    data.reset_index(inplace=True)

    grouped = data.groupby(["cohort_group", Jahr_new])

    cohorts = grouped.agg({Kundennummer_new: pd.Series.nunique, Umsatz_new: np.sum})

    cohorts.rename(columns={Kundennummer_new: "total_customers", Umsatz_new: "total_revenue"}, inplace=True)

    def cohort_period(data_frame):
        data_frame["cohort_period"] = np.arange(len(data_frame)) + 1
        return data_frame

    cohorts = cohorts.groupby(level=0).apply(cohort_period)

    cohorts.reset_index(inplace=True)

    pivot_anzahl = cohorts.pivot(index="cohort_group", columns=Jahr_new, values='total_customers')
    pivot_cohorts = cohorts.pivot(index="cohort_group", columns=Jahr_new, values='total_revenue')

    # with pd.ExcelWriter('Cohort_final.xlsx') as writer:  # doctest: +SKIP
    # data1.to_excel(writer, sheet_name='Cohort_raw', index= False)
    # cohorts.to_excel(writer, sheet_name='Cohort_group_total', index= False)
    # pivot_cohorts.to_excel(writer, sheet_name='Cohort_Volume', index= True)
    # pivot_anzahl.to_excel(writer, sheet_name='Cohort_Number', index= True)

    #########################################################################################
    writer = ExcelWriter('Cohort_final.xlsx')
    data1.reset_index(inplace=True)
    data1.to_excel(writer, sheet_name='Cohort_raw', index=False, encoding="utf-8")
    cohorts.to_excel(writer, sheet_name='Cohort_group_total', index=False, encoding="utf-8")
    pivot_cohorts.to_excel(writer, sheet_name='Cohort_Volume_all', index=True, encoding="utf-8")
    pivot_anzahl.to_excel(writer, sheet_name='Cohort_Number_all', index=True, encoding="utf-8")

    if len(txt4.get()) != 0:
        Umsatzart = int(txt4.get()) - 1
        #Umsatzart = Umsatzart - 1
        Umsatzart_new = data.columns[Umsatzart]
        Umsatzart_list = data[Umsatzart_new].unique().tolist()
        Umsatzart_list_length = len(data[Umsatzart_new].unique().tolist())
        #data_umsatzart1[Umsatzart_new].fillna("No_data", inplace=True)
        for i in Umsatzart_list:
            data_umsatzart = data_umsatzart1.copy()
            data_umsatzart = data_umsatzart[data_umsatzart[Umsatzart_new] == i]
            data_umsatzart.set_index(Kundennummer_new, inplace=True)
            data_umsatzart["cohort_group"] = data_umsatzart.groupby(level=0)[Jahr_new].min()
            data_umsatzart.reset_index(inplace=True)
            ###############################################################
            data_umsatzart.to_excel(writer, sheet_name=str(i) + "-r", encoding="utf-8", index=False)
            ###############################################################
            grouped_umsatzart = data_umsatzart.groupby(["cohort_group", Jahr_new])
            cohorts_umsatzart = grouped_umsatzart.agg({Kundennummer_new: pd.Series.nunique, Umsatz_new: np.sum})
            cohorts_umsatzart.rename(columns={Kundennummer_new: "total_customers", Umsatz_new: "total_revenue"},
                                     inplace=True)

            def cohort_period(data_frame):
                data_frame["cohort_period"] = np.arange(len(data_frame)) + 1
                return data_frame

            cohorts_umsatzart = cohorts_umsatzart.groupby(level=0).apply(cohort_period)
            cohorts_umsatzart.reset_index(inplace=True)
            pivot_anzahl_umsatzart = cohorts_umsatzart.pivot(index="cohort_group", columns=Jahr_new,
                                                             values='total_customers')
            pivot_cohorts_umsatzart = cohorts_umsatzart.pivot(index="cohort_group", columns=Jahr_new,
                                                              values='total_revenue')

            # with pd.ExcelWriter('Cohort_final.xlsx') as writer:  # doctest: +SKIP
            #     pivot_cohorts_umsatzart.to_excel(writer, sheet_name=i, index= True)
            #     pivot_anzahl_umsatzart.to_excel(writer, sheet_name=i, index= True)

            pivot_anzahl_umsatzart.to_excel(writer, sheet_name=str(i) + "_n", encoding="utf-8")
            pivot_cohorts_umsatzart.to_excel(writer, sheet_name=str(i) + "_sum", encoding="utf-8")

    writer.save()

    #########################################################################################
    messagebox.showinfo('Progress Update', 'Done!')


###################################################
###################################################
def browse_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global folder_path
    folder_path = askopenfilename()


#######################################################

btn1 = Button(window, text="Choose File", command=browse_button)
btn2 = Button(window, text="Start process !", bg="green", command=clicked)
##btn3 = Button(window, text='Set directory', command=set_dir)

btn1.grid(column=1, row=6)
btn2.grid(column=1, row=7)
##btn3.grid(column = 1, row=10)

window.mainloop()
