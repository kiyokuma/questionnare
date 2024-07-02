# coding: utf-8
import os
import pathlib as pl
import csv
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox as messagebox
import openpyxl
    
def csv_cp932_read(path_name):
    csv_header = ["code","name"]
    csv_data = {}
    with open(path_name ,'r',encoding="cp932") as f:
        for row in csv.DictReader(f, csv_header):
            csv_data[row['code']] = row['name']
    return csv_data

def csv_shiftjis_read(path_name):
    csv_header = ["code","name"]
    csv_data = {}
    with open(path_name ,'r',encoding="shift_jis") as f:
        for row in csv.DictReader(f, csv_header):
            csv_data[row['code']] = row['name']
    return csv_data

def file_shiftjis_read(input_path):
    #ファイルの呼び出し
    typ = [('', '*.csv')] 
    root_small = tk.Tk()
    fle = filedialog.askopenfilename(filetypes = typ, initialdir = input_path)
    if fle == "":
        fle = None
        messagebox.showwarning("エラー","ファイルの選択に失敗しました。")
        root_small.destroy()
        return
    else:
        fle_read = pd.read_csv(fle,encoding='shift_jis')
        root_small.destroy()
        return fle_read

def file_cp932_read(input_path):
    #ファイルの呼び出し
    typ = [('', '*.csv')]
    root_small = tk.Tk()
    fle = filedialog.askopenfilename(filetypes = typ, initialdir = input_path)
    if fle == "":
        fle = None
        messagebox.showwarning("エラー","ファイルの選択に失敗しました。")
        root_small.destroy()
        return
    else:
        fle_read = pd.read_csv(fle,encoding='cp932')
        root_small.destroy()
    return fle_read

def questionnaire_xlsx_read(input_path):
    #ファイルの呼び出し
    typ = [('', '*.xlsx')]
    root_small = tk.Tk()
    fle = filedialog.askopenfilename(filetypes = typ, initialdir = input_path)
    if fle == "":
        fle = None
        messagebox.showwarning("エラー","ファイルの選択に失敗しました。")
        root_small.destroy()
        return
    else:
        fle_read = pd.read_excel(fle, engine='openpyxl',index_col=0,header=2)
        root_small.destroy()
    return fle_read

def file_tokenized_read(input_path):
    pd.options.display.max_rows = None
    pd.options.display.max_columns = None
    #ファイルの呼び出し
    typ = [('', '*.csv')]
    root_small = tk.Tk()
    fle = filedialog.askopenfilename(filetypes = typ, initialdir = input_path)
    if fle == "":
        fle = None
        messagebox.showwarning("エラー","ファイルの選択に失敗しました。")
        root_small.destroy()
        return
    else:
        root_small.destroy()
        df = pd.DataFrame()
        with open(fle, "r", encoding="cp932", errors="", newline="" ) as f:
            lst = csv.reader(f, delimiter=",")
            df = pd.DataFrame(lst)
        return df

def message_read_complete():
    messagebox.showinfo(
        title = "message",
        message = "ファイルの取り込み",
        detail = "取り込みが完了しました。")        
def message_save_complete():
    messagebox.showinfo(
        title = "message",
        message = "ファイルの加工",
        detail = "加工が完了しました。")