# -*- coding: utf-8 -*-
"""
Created on Wed Mar  9 22:56:35 2022

@author: yo
"""
import os
import re
import xml.etree.ElementTree as ET
import xlsxwriter
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk

def selectPath(flag):
    hint['text'] = ""
    win.geometry(f'{width}x{height}')
    _path = filedialog.askdirectory(title='Select directory')
    if not flag:
        filepath.set(_path)
        file_entry.focus()
        file_entry.icursor(len(_path))
    else:
        savepath.set(_path)
        save_entry.focus()
        save_entry.icursor(len(_path))

def selectFile():
    hint['text'] = ""
    win.geometry(f'{width}x{height}')
    _path = filedialog.askopenfilename(title='Select file',
                                       filetypes=[("xml files", "*.xml")])
    filepath.set(_path)
    file_entry.focus()
    file_entry.icursor(len(_path)) 
    
def XMLToExcel(_filepath, savepath):
    filename = re.compile(r'.*\/(.*?)\.xml').search(_filepath).group(1)
    tree = ET.parse(_filepath)
    root = tree.getroot()
    savepath += "/" + filename
    if os.path.isfile(savepath + ".xlsx"):
        i = 1
        while True:
            if not os.path.isfile(f'{savepath} ({i}).xlsx'):
                savepath = f'{savepath} ({i}).xlsx'
                break
            else:
                i += 1
    else:
        savepath += ".xlsx"
        
    workbook = xlsxwriter.Workbook(savepath)
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, "section")
    worksheet.write(0, 1, "entry")
    row = 1
    for child in root.findall(".//mainsection[@title='Summary']")[0]:
        col = 0
        worksheet.write(row, col, child.get('title'))
        col += 1
        for item in child:
            worksheet.write(row, col, item.get('title'))
            col += 1
        row += 1
    workbook.close()
    return savepath

def pathHandle(_path, savepath):
    if _path == "" or savepath == "":
        win.geometry(f'{width}x{height}')
        hint['text'] = "請選擇資料夾/檔案"
        return
    elif ".xml" in _path:
        output_file = XMLToExcel(_path, savepath)
        if len(output_file) > 117:
            expend = int((len(output_file)+12)/40) + 1
            expend = height + (expend-3)*12
            win.geometry(f'{width}x{expend}')
        hint['text'] = f"完成，已輸出至{output_file}"
    else:
        dirpath = re.compile(r'.*(\/.*)').search(_path).group(1)
        dirpath += "_excel"
        if os.path.isdir(savepath + dirpath):
            i = 1
            while True:
                if not os.path.isdir(f'{savepath}{dirpath} ({i})'):
                    savepath = f'{savepath}{dirpath} ({i})'
                    break
                else:
                    i += 1
        else:
            savepath += dirpath
            
        dirs = os.listdir(_path)
        handle_files = []
        for file_name in dirs:
            if ".xml" in file_name:
                handle_files.append(file_name)
        if len(handle_files) == 0:
            win.geometry(f'{width}x{height}')
            hint['text'] = "此資料夾中無xml檔"
            return
        else:
            os.mkdir(savepath)
            hint.grid_forget()
            progressbar.grid(column=0,row=2)
            finished_files = 0
            for file_name in handle_files:
                file_path = _path + "/" + file_name
                XMLToExcel(file_path, savepath)
                finished_files += 1
                progressbar["value"] = int(finished_files/len(handle_files)*100)
                win.update()
            progressbar.grid_forget()
            hint.grid(column=0,row=2)
            if len(savepath) > 117:
                expend = int((len(savepath)+12)/40) + 1
                expend = height + (expend-3)*12
                win.geometry(f'{width}x{expend}')
            hint['text'] = f"完成，已輸出至{savepath}"


win = tk.Tk()  
filepath = tk.StringVar()
savepath = tk.StringVar()
width = 500
height = 150
x = int((win.winfo_screenwidth()-width)/2)
y = int((win.winfo_screenheight()-height)/2)

win.title("XML to excel")
win.geometry(f'{width}x{height}+{x}+{y}')
win.resizable(0, 0)
file_entry = ttk.Entry(win, textvariable=filepath, width=40)
ttk.Button(win, text="選取資料夾...", command=lambda:selectPath(0)).grid(column=1,row=0, pady=10)
ttk.Button(win, text="選取檔案...", command=lambda:selectFile()).grid(column=2,row=0, pady=10)
save_entry = ttk.Entry(win, textvariable=savepath, width=40)
ttk.Button(win, text="選取儲存路徑...", command=lambda:selectPath(1)).grid(column=1,row=1, pady=10, columnspan=2, sticky=tk.E+tk.W)
hint = ttk.Label(win, text="", wraplength=300)
progressbar = ttk.Progressbar(win, length=285, mode="determinate")
ttk.Button(win, text="開始", command=lambda:pathHandle(filepath.get(), savepath.get())).grid(column=1,row=2, pady=10, columnspan=2, sticky=tk.E+tk.W)

file_entry.grid(column=0,row=0, padx=10)
save_entry.grid(column=0,row=1)
hint.grid(column=0,row=2)
progressbar.grid(column=0,row=2)
progressbar.grid_forget()
win.mainloop()
