# -*- coding: utf-8 -*-
"""
Created on Sat Mar 20 08:26:16 2021

@author: Sean Moore
"""

import pandas as pd
import yfinance as yf
import tkinter as tk
from tkinter import messagebox
import tkinter.filedialog
import os

#A function that will get and return the correlation matrix for the given tickers.
def get_correlation_matrix(tickers, period='5y', interval='1d', cor_method='pearson'): 
    df = pd.DataFrame()
    for t in tickers: 
        tmp_ticker = yf.Ticker(t)
        tmp_hist = tmp_ticker.history(period=period, interval=interval)
        df[t] = tmp_hist['Open']
    return df.corr(method=cor_method)

#display(get_correlation_matrix(['MSFT', 'AAPL', 'TSLA', 'CROX', 'GME', 'AMC'])) #For testing.
    
#A function that will run this application with the input. 
def punch_it(): 
    #Get the list of tickers. 
    tickers = ticker_entry.get()
    if len(tickers) == 0: 
        messagebox.showerror("Error", "It appears that you entered no tickers.")
        return
    tickers = [x.strip() for x in tickers.split(',')]
    
    path =  tk.filedialog.askdirectory()
    filename = "correlationMatrix.xlsx" 
    file_path = os.path.join(path, filename)
    
    #get the correlation matrix. 
    try: 
        cm = get_correlation_matrix(tickers)
    except Exception: 
        messagebox.showerror("Error", "There was an error computing the correlation matrix.")
        return
    
    #Save the resulting data frame to an excel file. 
    try: 
        cm.to_excel(file_path)
    except Exception: 
        messagebox.showerror("Error", "There was an error saving the excel file.")
        return
    
    #Close the applicatoin when we are done. 
    tk.messagebox.showinfo(title="Done.", message="The excel file has been saved. The program will now exit.")
    application_window.destroy()
    
#The application. 
application_window = tk.Tk() 
application_window.title("Correlation Matrix solver")

ticker_input_frame = tk.Frame(borderwidth=10)
input_label = tk.Label(text="Tickers:", master=ticker_input_frame)
input_label.grid(row=0, column=0)
ticker_entry = tk.Entry(master=ticker_input_frame, width=100)
ticker_entry.grid(row=0, column=1)
ticker_input_frame.grid(row=0, column=0)

#The button that will begin the process and save the result. 
button_frame = tk.Frame(borderwidth=10)
go_button = tk.Button(master=button_frame, text="Punch it!", command=punch_it)
go_button.pack()
button_frame.grid(row=1, column=0)

#Start the main application loop. 
application_window.mainloop()