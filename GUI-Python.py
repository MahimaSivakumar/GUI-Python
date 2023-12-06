
"""
@author: mahimasivakumar
"""
# Packages
import pandas as pd
import tkinter as tk
import seaborn as sb
from pandastable import Table, TableModel
from tkinter import ttk
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
import matplotlib.pyplot as plt
import numpy as np
from tkinter import messagebox
import pickle


#read the raw data from excel
try:
    source = pd.read_excel('EnergyCallcentre.xlsx')
except Exception as e:
    print(" Try Reading the Excel from the Correct File Path")

#sampling of raw data 
samp=source.sample(100, random_state=2583) 
#sorting based on index       
sample=samp.sort_index()

#Converting column value into input list
month = sample['Month'].unique().tolist()
vht = sample['VHT'].unique().tolist()
tod =sample['ToD'].unique().tolist()

#Aggregation of the Data
calls = sample.groupby(by=['Month','VHT','ToD'],as_index=False)['CallsAbandoned'].mean()
stats = sample.groupby(['Month' , 'ToD'], as_index=False).agg(['min', 'max', 'mean']).reset_index()


#finding correlation between variables 
Correlation = sample.corr()
print(Correlation)

#log file 
def logfile():
    
    with open('logpickle.pkl', 'wb') as f:
        pickle.dump(GUIscreen ,f)
        
    with open('logpickle.pkl', 'wb') as f:
         log = pickle.loads(f)
        
    return log
 


#GUI Screen Function
class CallGUI():
    
    #main GUI Screen
    def __init__(self,mains,sample,month,vht,calls,tod,stats):
        self.mains=mains     
        self.sample=sample
        self.month = month
        self.vht = vht
        self.calls = calls
        self.tod = tod
        mains.title(" Energy Call Center ")
        
        # Getting Input from the User
        mains.label = tk.Label(mains, text = " Enter the Month " , font = ("Times New Roman" , 18)).place(x=350,y=80)
        varmonth = tk.StringVar()
        varmonth.set('Select the Appropriate Month')
        dropdown = tk.OptionMenu(mains,varmonth,*month).place(x=620,y=80)
        mains.todlabel = tk.Label(mains, text = " Enter the Time of Day" , font = ("Times New Roman" , 18)).place(x=350,y=120)
        vartod = tk.StringVar(mains)
        vartod.set('Select the Time of Day')
        droptod = tk.OptionMenu(mains,vartod,*tod).place(x=620,y=120)
        mains.vht = tk.Label(mains,text ="Virtual Hold Technology",font = ("Times New Roman" , 18)).place(x=350 , y=30)
        varvht = tk.StringVar()
        varvht.set('')
        for item in vht:
            button = tk.Radiobutton(mains, text=item , variable=varvht, value=item).pack(ipadx =3 , ipady= 5)

        #filter the table based on user input
        def filters():
            
            if (varmonth.get()=='' or vartod.get()=='' or varvht.get()==''):
                messagebox.showwarning("showwarning", "Please Enter all the details")
                
            else:
                filters = tk.Toplevel(self.mains)
                filters.title("Filtered data based on user Input")
                mon = sample.loc[sample["Month"]== varmonth.get()]
                tod = mon.loc[mon["ToD"]== vartod.get()]
                VHT = tod.loc[tod["VHT"]== varvht.get()]
                table=Table(filters,dataframe=VHT)
                table.show()
            
        apply = tk.Button(mains, text = "Filter the Data" ,command=filters,height=2,width=9).place(x=950,y=80)
        frame = ttk.Frame(mains)
        frame.place(relx=0.1,rely=0.3)
        mains.mainlbl = tk.Label(mains, text = "Data Exploration" , font =("Times New Roman",22)).place(x=650 , y=170)
        
        #Data Exploration Options
        mains.vhtlabl = tk.Label(mains,text = "Efficiency of the VHT " , fg='black',font =("Times New Roman",18)).place(x=350,y=220)
        vht_btn = tk.Button(mains, text = "Find Out", command = self.vht_start,height=2,width=9).place(x=850,y=220 ,height=40)
        mains.labl = tk.Label(mains, text = "Summary Statistics of Call Center " , fg='black', font=("Times New Roman", 18)).place(x=350,y=260)
        btn = tk.Button(mains, text = "Click Here" ,  command=self.start, height=2,width=9).place(x=850,y=260)
        mains.freq = tk.Label(mains, text = " Call Frequency of the Data " , fg='black', font=("Times New Roman", 18)).place(x=350,y=320)
        freqbtn = tk.Button(mains, text = "Get Data " ,  command=self.frequency_call, height=2,width=9).place(x=850,y=320)
        #logbtn = tk.Button(mains, text = "View Log File" ,command= method, height=2,width=9).place(x=850,y=380)
        
   #new window to display the efficiency of the VHT 
    def vht_start(self):
        self.newwindow_vht()   
        
        #Extracting sample data from source and taking list of month and VHT for UI
    def start(self): 
        self.newwindow()
        
        
       # To Open New Window  for summary statistucs
    def newwindow(self):
        windows = tk.Toplevel(self.mains)
        windows.geometry("1820x1080")
        windows.title("Activity and Performance")
        frame = ttk.Frame(windows)
        frame.place(relx=0.1,rely=0.3)
        
        #to display the table and graph of summary statistics
        def call_summary():
            mon=var.get()
            tod=vare.get()
            summ_month = stats.loc[stats["Month"]== mon]
            summ_final = summ_month.loc[summ_month['ToD']==tod]
            table=Table(frame,dataframe=summ_final)
            table.show()
            fig=plt.Figure(figsize=(3,1),dpi=50)
            ax1= fig.add_subplot(111)
            sb.boxplot(data = sample, ax=ax1)
            bar1=FigureCanvasTkAgg(fig,windows)
            bar1.draw()
            bar1.get_tk_widget().place(relx=0.55,rely=0.2 ,height=550,width=650)
            
        #To display Month Dropdown in UI
        windows.lbl = tk.Label(windows, text = " Enter the Month" , font = ("Times New Roman" , 18)).place(x=280,y=20)
        var = tk.StringVar(windows)
        var.set('Select the Appropriate Month')
        dropdown = tk.OptionMenu(windows,var,*month).pack(padx=10,pady=30)
        
        #To display ToD Dropdown in UI
        windows.tod = tk.Label(windows, text = " Enter the Time of Day" , font = ("Times New Roman" , 18)).place(x=280,y=50)
        vare = tk.StringVar(windows)
        vare.set('Select the Time of Day')
        droptod = tk.OptionMenu(windows,vare,*tod).place(x=650,y=60)
        get_button = tk.Button(windows, text='Get Summary Data' , command=call_summary ).pack()

    
    #New Window Functionality for VHT Efficiency       
    def newwindow_vht(self):
        vht_windows = tk.Toplevel(self.mains)
        vht_windows.geometry("1820x1080")
        vht_windows.title("Performance level of VHT")
        frame = ttk.Frame(vht_windows)
        frame.place(relx=0.1,rely=0.3)
        
        #to display table and graph of VHT
        def display_selected(inputs):
            inputs = var.get()
            months = calls.loc[calls["Month"]== inputs]
            df = months
            table = Table(frame, dataframe=df)
            table.show()
            fig=plt.Figure(figsize=(3,1),dpi=50)
            ax1= fig.add_subplot(111)
            sb.barplot(x = df.VHT, y = df.CallsAbandoned, data = df, hue = df.ToD , ax=ax1)
            bar1=FigureCanvasTkAgg(fig,vht_windows)
            bar1.draw()
            bar1.get_tk_widget().place(relx=0.55,rely=0.2 ,height=550,width=550)
                      
        #To display Month Dropdown in UI
        vht_windows.lbl = tk.Label(vht_windows, text = " Enter the Month ( Calculate Average No.of CallsAbandoned)" , font = ("Times New Roman" , 18)).place(x=350,y=60)
        var = tk.StringVar()
        var.set('Select the Appropriate Month')
        dropdown = tk.OptionMenu(vht_windows,var,*month,command=display_selected).place(x=850 , y=60)
    
    #CAll frequency function
    def frequency_call(self):
        callfreq = sample.groupby(by=['Month','ToD'],as_index=False)['CallsOffered'].mean()
        freq=tk.Toplevel(self.mains)
        freq.title("Call Frequency")
        
        # Graph display of call frequency
        def freqmonth(month):
            month=var.get()
            finalmonth = callfreq.loc[callfreq["Month"]== month]
            fig=plt.Figure(figsize=(3,1),dpi=50)
            ax1= fig.add_subplot(111)
            sb.barplot(x = 'ToD', y = 'CallsOffered', data = finalmonth, ax=ax1)
            bar1=FigureCanvasTkAgg(fig,freq)
            bar1.draw()
            bar1.get_tk_widget().place(relx=0.3,rely=0.2 ,height=550,width=950)
        
        #To display Month Dropdown in UI
        freq.lbl = tk.Label(freq, text = " Enter the Month" , font = ("Times New Roman" , 18)).place(x=550,y=20)
        var = tk.StringVar()
        var.set('Select the Appropriate Month')
        dropdown = tk.OptionMenu(freq,var,*month , command = freqmonth).place(x=750 , y=20)
        

           
#Initiation to call GUI screen     
root =tk.Tk()
GUIscreen = CallGUI(root,sample,month,vht,calls,tod,stats)
root.geometry("1820x1080")
root.mainloop()


