import tkinter as tk
from tkinter import *
from tkinter import ttk, font
import os
import pandas as pd
import datetime as dt
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import win32com.client as win32
# import requests


def loaddf():
    global df
    df = pd.read_csv('masterdata.csv', index_col=0)
    df['Tmp'] = df.index.astype(int)

def savedf():
    df.to_csv('masterdata.csv')

def loadqcsdf():
    global qcsdf
    qcsdf = pd.read_csv('masterdataqcs.csv', index_col=0)
    qcsdf['Tmp'] = qcsdf.index.astype(int)

def saveqcsdf():
    qcsdf.to_csv('masterdataqcs.csv')

def loadpsadf():
    global psadf
    psadf = pd.read_csv('masterdatapsa.csv', index_col=0)
    psadf['Tmp'] = psadf.index.astype(int)
    psadf = psadf.astype(str)

def savepsadf():
    psadf.to_csv('masterdatapsa.csv')

def drawbuttons():
    Button(main_screen, text="People", command=lambda: [subgroup.set("People"), loadframe()], bg=btcolor, relief=FLAT).place(relx=offset, rely=0.05, relheight=buttonrelheight, relwidth=buttonrelwidth)
    if 'QAT' in dropselect.get():

        Button(main_screen, text="Triage", command=lambda: [subgroup.set("Triage"), loadframe()], bg=btcolor,
               relief=FLAT).place(relx=offset + buttonrelwidth * 2, rely=0.05, relheight=buttonrelheight,
                                  relwidth=buttonrelwidth)
    else:
        Button(main_screen, text="Safety", command=lambda: [subgroup.set("Safety"), loadframe()], bg=btcolor, relief=FLAT).place(relx=offset+buttonrelwidth*2, rely=0.05, relheight=buttonrelheight, relwidth=buttonrelwidth)

    if 'QAT' in dropselect.get():
        Button(main_screen, text="Deviations", command=lambda: [subgroup.set("Deviations"), loadframe()], bg=btcolor,
               relief=FLAT).place(relx=offset + buttonrelwidth * 4, rely=0.05, relheight=buttonrelheight,
                                  relwidth=buttonrelwidth)
    else:
        Button(main_screen, text="Quality", command=lambda: [subgroup.set("Quality"), loadframe()], bg=btcolor, relief=FLAT).place(relx=offset+buttonrelwidth*4, rely=0.05, relheight=buttonrelheight, relwidth=buttonrelwidth)

    if 'QAT' in dropselect.get():
        Button(main_screen, text="CAPA", command=lambda: [subgroup.set("CAPA"), loadframe()], bg=btcolor, relief=FLAT).place(relx=offset+buttonrelwidth*6, rely=0.05, relheight=buttonrelheight, relwidth=buttonrelwidth)
    else:
        Button(main_screen, text="Delivery", command=lambda: [subgroup.set("Delivery"), loadframe()], bg=btcolor,
               relief=FLAT).place(relx=offset+buttonrelwidth*6, rely=0.05, relheight=buttonrelheight, relwidth=buttonrelwidth)


    if 'QAT' in dropselect.get():
        Button(main_screen, text="BR", command=lambda: [subgroup.set("BR"), loadframe()], bg=btcolor, relief=FLAT).place(relx=offset+buttonrelwidth*8, rely=0.05, relheight=buttonrelheight, relwidth=buttonrelwidth)
    else:
        Button(main_screen, text="Cost", command=lambda: [subgroup.set("Cost"), loadframe()], bg=btcolor,
           relief=FLAT).place(relx=offset+buttonrelwidth*8, rely=0.05, relheight=buttonrelheight, relwidth=buttonrelwidth)

def loadframe():

    def drawdayframe():
        dayframe = LabelFrame(labelframe)
        # dayframe.grid(row=32, column=2, sticky=N + S + E + W)
        dayframe.grid(row=32, column=1, sticky=N + S + E + W)
        dayframe.place(relx=0.005, rely=0, relheight=.995, relwidth=0.07)

        for r in range(32):
            Grid.rowconfigure(dayframe, r, weight=1)
            for c in range(1):
                Grid.columnconfigure(dayframe, c, weight=1)
                if r == 0 and c == 0:
                    Label(dayframe, text='Day', bg=bgcolor).grid(row=r, column=c, sticky=N + S + E + W)
                #
                #     elif r == 0 and c == 1:
                #         Label(dayframe, text='Night',bg ="gray80").grid(row=r, column=c, sticky=N+S+E+W)
                else:
                    if str(f"{r:0>2}" + ' ' + str(date.strftime("%b")) + ' ' + str(date.strftime("%y"))) in df[
                        (df.Tier.astype('str') == dropselect.get()) & (df.Category == subgroup.get()) & (
                                subgroup.get() != 'People') & (pd.DatetimeIndex(df.Date).month == date.month) & (
                                pd.DatetimeIndex(df.Date).year == date.year)].Date.to_string():
                        Label(dayframe, text='%s' % (r), bg="orange red").grid(row=r, column=c, sticky=N + S + E + W)
                    elif str(f"{r:0>2}" + ' ' + str(date.strftime("%b")) + ' ' + str(date.strftime("%y"))) in df[
                        (df.Tier.astype('str') == dropselect.get()) & (df.Category == subgroup.get()) & (
                                subgroup.get() == 'People') & (pd.DatetimeIndex(df.Date).month == date.month) & (
                                pd.DatetimeIndex(df.Date).year == date.year)].Date.to_string():
                        Label(dayframe, text='%s' % (r), bg="green").grid(row=r, column=c, sticky=N + S + E + W)
                    elif r < int(date.strftime("%d")):
                        Label(dayframe, text='%s' % (r), bg="pale green").grid(row=r, column=c, sticky=N + S + E + W)
                    else:
                        Label(dayframe, text='%s' % (r)).grid(row=r, column=c, sticky=N + S + E + W)

    def drawtreeview():
        def item_selected(event):
            def savebutton(dtext, stext, indexD):
                df.at[indexD, 'Description'] = dtext.get("1.0", 'end-1c')
                df.at[indexD, 'Status'] = stext.get("1.0", 'end-1c')
                savedf()
                drawdayframe()
                drawtreeview()
                popframe.quit()
                popframe.destroy()

            def delbutton(indexD):
                # print(df)
                df.drop(int(float(indexD)), inplace=True)
                savedf()
                loadframe()
                popframe.quit()
                popframe.destroy()

            def exitbutton():
                popframe.quit()
                popframe.destroy()

            item_id = event.widget.focus()
            item = event.widget.item(item_id)
            values = item['values']
            dateD = values[0]
            whereD = values[1]
            descriptionD = values[2]
            statusD = values[3]
            indexD = values[4]

            popframe = Toplevel(main_screen)
            popframe.geometry("800x600")
            popframe.title('Event: ' + descriptionD)
            datewhereframe = Label(popframe,
                                   text='Date: ' + dateD + '   Raised by: ' + whereD + '    Index: ' + str(indexD))
            datewhereframe.place(relx=0, rely=0, relheight=.1, relwidth=1)

            descriptionframe = LabelFrame(popframe, text='Description')
            descriptionframe.place(relx=0, rely=0.1, relheight=.8, relwidth=.5)
            dtext = tk.Text(descriptionframe)
            dtext.insert(INSERT, descriptionD)
            dtext.place(relx=0, rely=0, relheight=1, relwidth=1)

            statusframe = LabelFrame(popframe, text='Status')
            statusframe.place(relx=.5, rely=0.1, relheight=.8, relwidth=.5)
            stext = tk.Text(statusframe)
            stext.insert(INSERT, statusD)
            stext.place(relx=0, rely=0, relheight=1, relwidth=1)

            Button(popframe, text="Save", command=lambda: savebutton(dtext, stext, indexD)).place(relx=0.0, rely=0.93,
                                                                                                  relheight=buttonrelheight,
                                                                                                  relwidth=.3)
            Button(popframe, text="Delete", command=lambda: delbutton(indexD)).place(relx=0.5 - 0.15, rely=0.93,
                                                                                     relheight=buttonrelheight,
                                                                                     relwidth=.3)
            Button(popframe, text="Close", command=exitbutton).place(relx=0.7, rely=0.93, relheight=buttonrelheight,
                                                                     relwidth=.3)

            popframe.mainloop()

        def new_entry_button():
            rbselection = StringVar()

            def savebutton(dtext, stext, rbselection):
                loaddf()
                indexD = df.index.max() + 1
                df.at[indexD, 'Description'] = dtext.get("1.0", 'end-1c')
                df.at[indexD, 'Status'] = stext.get("1.0", 'end-1c')
                df.at[indexD, 'Date'] = tdate
                df.at[indexD, 'Category'] = subgroup.get()
                df.at[indexD, 'Tier'] = dropselect.get()
                df.at[indexD, 'Tmp'] = indexD
                df.at[indexD, 'Raised by'] = rbselection.get()
                savedf()
                drawdayframe()
                drawtreeview()
                popframe.quit()
                popframe.destroy()

            def exitbutton():
                popframe.quit()
                popframe.destroy()

            popframe = Toplevel(main_screen)
            popframe.geometry("800x600")
            popframe.title('Add new event')

            rbselection.set('QA')
            options_list = ['QA', 'SE', 'TS', 'USP', 'DSP1', 'DSP2', 'SHF']

            rbselectlabel = Label(popframe, text='Raised By : ')
            rbselectlabel.place(relx=0.5 - .15, rely=0.01, relheight=.1, relwidth=0.2)
            rbselect = OptionMenu(popframe, rbselection, *options_list)
            rbselect.config(bg=bgcolor, relief=FLAT)
            rbselect.place(relx=.5, rely=0.03, relheight=buttonrelheight, relwidth=.15)

            descriptionframe = LabelFrame(popframe, text='Description')
            descriptionframe.place(relx=0, rely=0.1, relheight=.8, relwidth=.5)
            dtext = tk.Text(descriptionframe)
            dtext.insert(INSERT, 'New Description here')
            dtext.place(relx=0, rely=0, relheight=1, relwidth=1)

            statusframe = LabelFrame(popframe, text='Status')
            statusframe.place(relx=.5, rely=0.1, relheight=.8, relwidth=.5)
            stext = tk.Text(statusframe)
            stext.insert(INSERT, 'New Status here')
            stext.place(relx=0, rely=0, relheight=1, relwidth=1)

            Button(popframe, text="Save", command=lambda: savebutton(dtext, stext, rbselection)).place(relx=0.0,
                                                                                                       rely=0.93,
                                                                                                       relheight=buttonrelheight,
                                                                                                       relwidth=.3)
            Button(popframe, text="Close (without saving)", command=exitbutton).place(relx=0.7, rely=0.93,
                                                                                      relheight=buttonrelheight,
                                                                                      relwidth=.3)
            popframe.mainloop()

        tableframe = LabelFrame(labelframe)
        tableframe.place(relx=0.08, rely=0, relheight=0.4, relwidth=.84)

        columns = ('#1', '#2', '#3', '#4')
        tree = ttk.Treeview(tableframe, columns=columns, show='headings')
        tree.heading('#1', text='Date')
        tree.column('#1', width=50)
        tree.heading('#2', text='Raised By')
        tree.column('#2', width=50)
        tree.heading('#3', text='Description')
        tree.heading('#4', text='Status')
        tree.place(relx=0, rely=0, relheight=1, relwidth=.97)


        for index, row in df[(df['Tier'] == dropselect.get()) & (df['Category'] == subgroup.get())][["Date", "Raised by", "Description", "Status", "Tmp"]].iterrows():
                tree.insert('', 0, text=index, values=list(row))

        # scrollbar
        scrollbar = ttk.Scrollbar(tableframe, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.place(relx=0.97, rely=0, relheight=1, relwidth=1 - 0.97)
        Button(labelframe, text="Add new \nentry", command=lambda: new_entry_button(), bg=btcolor, relief=FLAT).place(
            relx=1 - 0.075, rely=0, relheight=0.4, relwidth=.07)
        tree.bind('<Double-Button-1>', item_selected)

    def drawmetrics():
            def create_plot():
                sns.set()

                pdf = pd.DataFrame(dict(Date=np.arange(50),
                                       Data=np.arange(50)-np.random.randn(50).cumsum()), columns=['Date', 'Data'])
                pdf2 = pd.DataFrame(dict(Date=np.arange(50),
                                        Data=np.arange(50)), columns=['Date', 'Data'])
                f, ax = plt.subplots(figsize=(10, 5))
                # sns.catplot(data=pdf, x='Date', y='Data', kind="bar" )
                sns.lineplot(data=pdf, x='Date', y='Data')
                sns.lineplot(data=pdf2, x='Date', y='Data')

                return f

            def tier2people():
                fig = create_plot()
                # canvas = FigureCanvasTkAgg(fig, master=metricframe)  # A tk.DrawingArea.
                # canvas.draw()
                # canvas.get_tk_widget().place(relx=0, rely=0, relheight=1, relwidth=.5)
                #
                # # toolbar = NavigationToolbar2Tk(canvas, metricframe)
                # # toolbar.update()
                #
                # fig2 = create_plot()
                # canvas2 = FigureCanvasTkAgg(fig2, master=metricframe)  # A tk.DrawingArea.
                # canvas2.draw()
                # canvas2.get_tk_widget().place(relx=0.5, rely=0, relheight=1, relwidth=.5)

            def tier2safety():
                fig = create_plot()
                canvas = FigureCanvasTkAgg(fig, master=metricframe)  # A tk.DrawingArea.
                canvas.draw()
                canvas.get_tk_widget().place(relx=0, rely=0, relheight=1, relwidth=.5)

                fig2 = create_plot()
                canvas2 = FigureCanvasTkAgg(fig2, master=metricframe)  # A tk.DrawingArea.
                canvas2.draw()
                canvas2.get_tk_widget().place(relx=0.5, rely=0, relheight=1, relwidth=.5)

            def tier2cost():
                def loadmetricframe():
                    if dropselect.get() == 'Prod. Yield':
                        def create_plot():
                            sns.set()

                            pdf = pd.DataFrame(dict(Date=np.arange(50),
                                                    Yield=np.arange(50) - np.random.randn(50).cumsum()),
                                               columns=['Date', 'Yield'])
                            pdf2 = pd.DataFrame(dict(Date=np.arange(50),
                                                     Yield=np.arange(50)), columns=['Date', 'Yield'])
                            f = plt.Figure(figsize=(6, 6))
                            ax = f.subplots()
                            sns.lineplot(data=pdf, x='Date', y='Yield', ax=ax)
                            sns.lineplot(data=pdf2, x='Date', y='Yield', ax=ax)
                            return f

                        yieldframe = LabelFrame(metricframe, text='Production Yield')
                        yieldframe.place(relx=0, rely=0.07, relheight=.9, relwidth=1)

                        fig = create_plot()
                        canvas = FigureCanvasTkAgg(fig, master=yieldframe)  # A tk.DrawingArea.
                        canvas.draw()
                        canvas.get_tk_widget().place(relx=0, rely=0, relheight=1, relwidth=.5)

                        fig2 = create_plot()
                        canvas2 = FigureCanvasTkAgg(fig2, master=yieldframe)  # A tk.DrawingArea.
                        canvas2.draw()
                        canvas2.get_tk_widget().place(relx=0.5, rely=0, relheight=1, relwidth=.5)

                dropselect = StringVar(metricframe)
                dropselect.set('Prod. Yield')
                options_list = ['Prod. Yield','Inventory Write Off','PCE Speed']

                Label(metricframe, text='Selection:').place(relx=0, rely=0)
                pageselect = OptionMenu(metricframe, dropselect, *options_list,
                                        command=lambda _: loadmetricframe())
                pageselect.config(bg=bgcolor, relief=FLAT)
                pageselect.place(relx=0.07, rely=0.0, relheight=0.07)
                loadmetricframe()

            def tier2quality():
                def loadmetricframe():
                    if dropselect.get() == 'Sample Submission':
                        def qcitem_selected(event):
                            def savebutton(dateD, areaD, processD, batchD, lineD, timeD, blockingD, indexD):
                                qcsdf.at[indexD, 'Date'] = dateD.get("1.0", 'end-1c')
                                qcsdf.at[indexD, 'Area'] = areaD.get()
                                qcsdf.at[indexD, 'Process'] = processD.get("1.0", 'end-1c')
                                qcsdf.at[indexD, 'Batch'] = batchD.get("1.0", 'end-1c')
                                qcsdf.at[indexD, 'Line'] = lineD.get()
                                qcsdf.at[indexD, 'Time'] = timeD.get("1.0", 'end-1c')
                                qcsdf.at[indexD, 'Blocking'] = blockingD.get()
                                saveqcsdf()
                                loadframe()
                                popframe.quit()
                                popframe.destroy()

                            def delbutton(indexD):
                                qcsdf.drop(indexD, inplace=True)
                                saveqcsdf()
                                loadframe()
                                popframe.quit()
                                popframe.destroy()

                            def exitbutton():
                                popframe.quit()
                                popframe.destroy()

                            item_id = event.widget.focus()
                            item = event.widget.item(item_id)
                            values = item['values']
                            dateD = values[0]
                            areaD = values[1]
                            processD = values[2]
                            batchD = values[3]
                            lineD = values[4]
                            timeD = values[5]
                            blockingD = values[6]
                            indexD = values[7]

                            popframe = Toplevel(main_screen)
                            popframe.geometry("720x240")
                            popframe.title(areaD +' '+ processD +' sample '+ batchD)


                            Label(popframe,text='Area: ').place(relx=0, rely=0, relheight=.1, relwidth=.1)
                            areasel = StringVar()
                            areasel.set(areaD)
                            areaselect = OptionMenu(popframe, areasel,'USP','DSP','DSP2','SHF')
                            areaselect.config(bg=bgcolor, relief=FLAT)
                            areaselect.place(relx=0.1, rely=0.0, relheight=buttonrelheight*2)

                            Label(popframe, text='Line: ').place(relx=0.4, rely=0, relheight=.1, relwidth=.1)
                            linesel = StringVar()
                            linesel.set(lineD)
                            lineselect = OptionMenu(popframe, linesel, '1', '2')
                            lineselect.config(bg=bgcolor, relief=FLAT)
                            lineselect.place(relx=0.5, rely=0.0, relheight=buttonrelheight*2)

                            Label(popframe, text='Blocking?: ').place(relx=0.8, rely=0, relheight=.1, relwidth=.1)
                            blockingsel = StringVar()
                            blockingsel.set(blockingD)
                            blockingselect = OptionMenu(popframe, blockingsel, 'Yes', 'No')
                            blockingselect.config(bg=bgcolor, relief=FLAT)
                            blockingselect.place(relx=.9, rely=0.0, relheight=buttonrelheight*2)


                            dateframe = LabelFrame(popframe, text='Date')
                            dateframe.place(relx=0, rely=0.1, relheight=.4, relwidth=.5)
                            datetext = tk.Text(dateframe)
                            datetext.insert(INSERT, dateD)
                            datetext.place(relx=0, rely=0, relheight=1, relwidth=1)

                            stepframe = LabelFrame(popframe, text='Process step/Type of Sample')
                            stepframe.place(relx=0, rely=0.5, relheight=.4, relwidth=.5)
                            steptext = tk.Text(stepframe)
                            steptext.insert(INSERT, processD)
                            steptext.place(relx=0, rely=0, relheight=1, relwidth=1)

                            batchframe = LabelFrame(popframe, text='Batch/Run')
                            batchframe.place(relx=.5, rely=0.1, relheight=.4, relwidth=.5)
                            batchtext = tk.Text(batchframe)
                            batchtext.insert(INSERT, batchD)
                            batchtext.place(relx=0, rely=0, relheight=1, relwidth=1)

                            timeframe = LabelFrame(popframe, text='Time')
                            timeframe.place(relx=.5, rely=0.5, relheight=.4, relwidth=.5)
                            timetext = tk.Text(timeframe)
                            timetext.insert(INSERT, timeD)
                            timetext.place(relx=0, rely=0, relheight=1, relwidth=1)


                            Button(popframe, text="Save", command=lambda: savebutton(datetext, areasel, steptext, batchtext, linesel, timetext, blockingsel, indexD), relief=FLAT).place(relx=0.0,rely=0.9,relheight=buttonrelheight*2,relwidth=.3)
                            Button(popframe, text="Delete", command=lambda: delbutton(indexD), relief=FLAT).place(relx=0.5 - 0.15, rely=0.9,relheight=buttonrelheight*2,relwidth=.3)
                            Button(popframe, text="Close", command=exitbutton, relief=FLAT).place(relx=0.7, rely=0.9,relheight=buttonrelheight*2, relwidth=.3)

                            popframe.mainloop()

                        def new_qc_entry_button():

                            def savebutton(dateD, areaD, processD, batchD, lineD, timeD, blockingD):
                                loadqcsdf()
                                indexD = qcsdf.index.max() + 1
                                # print(indexD)
                                qcsdf.at[indexD, 'Date'] = dateD.get("1.0", 'end-1c')
                                qcsdf.at[indexD, 'Area'] = areaD.get()
                                qcsdf.at[indexD, 'Process'] = processD.get("1.0", 'end-1c')
                                qcsdf.at[indexD, 'Batch'] = batchD.get("1.0", 'end-1c')
                                qcsdf.at[indexD, 'Line'] = lineD.get()
                                qcsdf.at[indexD, 'Time'] = timeD.get("1.0", 'end-1c')
                                qcsdf.at[indexD, 'Blocking'] = blockingD.get()
                                saveqcsdf()
                                loadframe()
                                popframe.quit()
                                popframe.destroy()

                            def exitbutton():
                                popframe.quit()
                                popframe.destroy()

                            popframe = Toplevel(main_screen)
                            popframe.geometry("720x240")
                            popframe.title('New QC sample submission')


                            Label(popframe,text='Area: ').place(relx=0, rely=0, relheight=.1, relwidth=.1)
                            areasel = StringVar()
                            areasel.set('USP')
                            areaselect = OptionMenu(popframe, areasel,'USP','DSP','DSP2','SHF')
                            areaselect.config(bg=bgcolor, relief=FLAT)
                            areaselect.place(relx=0.1, rely=0.0, relheight=buttonrelheight*2)

                            Label(popframe, text='Line: ').place(relx=0.4, rely=0, relheight=.1, relwidth=.1)
                            linesel = StringVar()
                            linesel.set('1')
                            lineselect = OptionMenu(popframe, linesel, '1', '2')
                            lineselect.config(bg=bgcolor, relief=FLAT)
                            lineselect.place(relx=0.5, rely=0.0, relheight=buttonrelheight*2)

                            Label(popframe, text='Blocking?: ').place(relx=0.8, rely=0, relheight=.1, relwidth=.1)
                            blockingsel = StringVar()
                            blockingsel.set('Yes')
                            blockingselect = OptionMenu(popframe, blockingsel, 'Yes', 'No')
                            blockingselect.config(bg=bgcolor, relief=FLAT)
                            blockingselect.place(relx=.9, rely=0.0, relheight=buttonrelheight*2)

                            dateframe = LabelFrame(popframe, text='Date')
                            dateframe.place(relx=0, rely=0.1, relheight=.4, relwidth=.5)
                            datetext = tk.Text(dateframe)
                            datetext.insert(INSERT, 'Date here')
                            datetext.place(relx=0, rely=0, relheight=1, relwidth=1)

                            stepframe = LabelFrame(popframe, text='Process step/Type of Sample')
                            stepframe.place(relx=0, rely=0.5, relheight=.4, relwidth=.5)
                            steptext = tk.Text(stepframe)
                            steptext.insert(INSERT, 'Process step/Type of Sample here')
                            steptext.place(relx=0, rely=0, relheight=1, relwidth=1)

                            batchframe = LabelFrame(popframe, text='Batch/Run')
                            batchframe.place(relx=.5, rely=0.1, relheight=.4, relwidth=.5)
                            batchtext = tk.Text(batchframe)
                            batchtext.insert(INSERT, 'Batch/Run here')
                            batchtext.place(relx=0, rely=0, relheight=1, relwidth=1)

                            timeframe = LabelFrame(popframe, text='Time')
                            timeframe.place(relx=.5, rely=0.5, relheight=.4, relwidth=.5)
                            timetext = tk.Text(timeframe)
                            timetext.insert(INSERT, 'Time here')
                            timetext.place(relx=0, rely=0, relheight=1, relwidth=1)

                            Button(popframe, text="Save", command=lambda: savebutton(datetext, areasel, steptext, batchtext, linesel, timetext, blockingsel), relief=FLAT).place(relx=0.0,rely=0.9,relheight=buttonrelheight*2,relwidth=.3)
                            Button(popframe, text="Close (without saving)", command=exitbutton, relief=FLAT).place(relx=0.7, rely=0.9,relheight=buttonrelheight*2, relwidth=.3)
                            popframe.mainloop()

                        loadqcsdf()

                        # Label(metricframe, text='QC Samples Schedule').place(relx=0, rely=0)
                        tableframe = LabelFrame(metricframe, text='QC Samples Schedule')
                        tableframe.place(relx=0, rely=0.07, relheight=.9, relwidth=.89)

                        columns = ('#1', '#2', '#3', '#4', '#5', '#6', '#7')
                        treeqc = ttk.Treeview(tableframe, columns=columns, show='headings')
                        treeqc.heading('#1', text='Date')
                        treeqc.column('#1', width=80)
                        treeqc.heading('#2', text='Area')
                        treeqc.column('#2', width=50)
                        treeqc.heading('#3', text='Process step/Type of Sample')
                        treeqc.heading('#4', text='Batch/Run')
                        treeqc.column('#4', width=100)
                        treeqc.heading('#5', text='Line')
                        treeqc.column('#5', width=50)
                        treeqc.heading('#6', text='Time of sample submission')
                        treeqc.column('#6', width=100)
                        treeqc.heading('#7', text='Blocking')
                        treeqc.column('#7', width=50)
                        treeqc.place(relx=0, rely=0, relheight=1, relwidth=1)

                        for index, row in qcsdf.iterrows():
                            treeqc.insert('', 0, text=index, values=list(row))

                        # scrollbar
                        scrollbary = ttk.Scrollbar(metricframe, orient=tk.VERTICAL, command=treeqc.yview)
                        treeqc.configure(yscroll=scrollbary.set)
                        scrollbary.place(relx=0.89, rely=0.1, relheight=.9, relwidth=.03)

                        scrollbarx = ttk.Scrollbar(metricframe, orient=tk.HORIZONTAL, command=treeqc.xview)
                        treeqc.configure(xscroll=scrollbarx.set)
                        scrollbarx.place(relx=0, rely=.95, relheight=.05, relwidth=.89)

                        Button(metricframe, text="Add new \nentry", command=lambda: new_qc_entry_button(), bg=btcolor,
                               relief=FLAT).place(relx=1 - 0.075, rely=0.1, relheight=.9, relwidth=.07)

                        treeqc.bind('<Double-Button-1>', qcitem_selected)

                dropselect = StringVar(metricframe)
                dropselect.set('Sample Submission')
                options_list = ['Sample Submission','Raw Material/Consumable Status','Deviations']

                Label(metricframe, text='Selection:').place(relx=0, rely=0)
                pageselect = OptionMenu(metricframe, dropselect, *options_list,
                                        command=lambda _: loadmetricframe())
                pageselect.config(bg=bgcolor, relief=FLAT)
                pageselect.place(relx=0.07, rely=0.0, relheight=0.07)
                loadmetricframe()

            def tier2delivery():

                def loadmetricframe():

                    if dropselect.get() == 'HVAC and Utilities':
                        def resultgen(input,target,r,c):
                            if input > target:
                                Label(hvframe, text='Pass', bg="pale green").grid(row=r, column=c, columnspan=1, sticky=NSEW)
                            else:
                                Label(hvframe, text='Fail', bg="orange red").grid(row=r, column=c, columnspan=1, sticky=NSEW)

                        def loadpidf():
                            global wfi21input
                            global wfi22input
                            global wfi23input
                            global tmacinput
                            global uspinput
                            global dspinput
                            global shfinput
                            global cominput

                            xl = win32.Dispatch('Excel.Application')
                            xl.Application.visible = False  # change to True if you are desired to make Excel visible

                            # disabled - you can use this to link any excel with datalink
                            # wb = xl.Workbooks.Open(os.path.abspath(os.getcwd() + '\\PIdatalinker.xlsx'))
                            # wb.Save()
                            # wb.Close()

                            pidf = pd.read_excel(os.getcwd() + '\\PIdatalinker.xlsx', sheet_name='Main')
                            # print(pidf['Time'].iloc[0])
                            wfi21input = Label(hvframe, text=pidf['Level'].iloc[0])
                            wfi21input.grid(row=2, column=7, columnspan=1, sticky=NSEW)
                            wfi22input = Label(hvframe, text=pidf['Level'].iloc[1])
                            wfi22input.grid(row=4, column=7, columnspan=1, sticky=NSEW)
                            wfi23input = Label(hvframe, text=pidf['Level'].iloc[2])
                            wfi23input.grid(row=6, column=7, columnspan=1, sticky=NSEW)
                            tmacinput = Label(hvframe, text=pidf['Level'].iloc[3])
                            tmacinput.grid(row=8, column=7, columnspan=1, sticky=NSEW)

                            resultgen(pidf['Level'].iloc[0], 20000, 2, 5)
                            resultgen(pidf['Level'].iloc[1], 20000, 4, 5)
                            resultgen(pidf['Level'].iloc[2], 20000, 6, 5)
                            resultgen(100-pidf['Level'].iloc[3], 50, 8, 5)

                            uspinput = Label(hvframe, text=pidf['Level'].iloc[4])
                            uspinput.grid(row=2, column=2, columnspan=1, sticky=NSEW)
                            dspinput = Label(hvframe, text=pidf['Level'].iloc[5])
                            dspinput.grid(row=4, column=2, columnspan=1, sticky=NSEW)
                            shfinput = Label(hvframe, text=pidf['Level'].iloc[6])
                            shfinput.grid(row=6, column=2, columnspan=1, sticky=NSEW)
                            cominput = Label(hvframe, text=pidf['Level'].iloc[7])
                            cominput.grid(row=8, column=2, columnspan=1, sticky=NSEW)

                            resultgen(pidf['Level'].iloc[4], 1000, 2, 0)
                            resultgen(pidf['Level'].iloc[5], 1100, 4, 0)
                            resultgen(pidf['Level'].iloc[6], 90, 6, 0)
                            resultgen(pidf['Level'].iloc[7], 600, 8, 0)

                        hvframe = LabelFrame(metricframe, text='HVAC and Utilities')
                        hvframe.place(relx=0, rely=0.07, relheight=.9, relwidth=1)
                        hvframe.grid_columnconfigure(0, weight=1)
                        hvframe.grid_columnconfigure(1, weight=1)
                        hvframe.grid_columnconfigure(2, weight=1)
                        hvframe.grid_columnconfigure(3, weight=1)
                        hvframe.grid_columnconfigure(4, weight=1)
                        hvframe.grid_columnconfigure(5, weight=1)
                        hvframe.grid_columnconfigure(6, weight=1)
                        hvframe.grid_columnconfigure(7, weight=1)
                        hvframe.grid_columnconfigure(8, weight=1)
                        hvframe.grid_rowconfigure(0, weight=1)
                        hvframe.grid_rowconfigure(1, weight=1)
                        hvframe.grid_rowconfigure(2, weight=1)
                        hvframe.grid_rowconfigure(3, weight=1)
                        hvframe.grid_rowconfigure(4, weight=1)
                        hvframe.grid_rowconfigure(5, weight=1)
                        hvframe.grid_rowconfigure(6, weight=1)
                        hvframe.grid_rowconfigure(7, weight=1)
                        hvframe.grid_rowconfigure(8, weight=1)
                        # usptbc = IntVar()
                        # dsptbc = IntVar()
                        # shftbc = IntVar()
                        # comtbc = IntVar()
                        # wfi21tbc = IntVar()
                        # wfi22tbc = IntVar()
                        # wfi23tbc = IntVar()
                        # tmactbc =IntVar()

                        # usptbc.trace_add('write',usptbcchanged)
                        # dsptbc.trace_add('write',dsptbcchanged)
                        # shftbc.trace_add('write', shftbcchanged)
                        # comtbc.trace_add('write',comtbcchanged)
                        # wfi21tbc.trace_add('write',wfi21tbcchanged)
                        # wfi22tbc.trace_add('write',wfi22tbcchanged)
                        # wfi23tbc.trace_add('write',wfi23tbcchanged)
                        # tmactbc.trace_add('write',tmactbcchanged)

                        Label(hvframe,text='HVAC',relief="ridge",borderwidth = 1).grid(row=0, column=0,columnspan=4,sticky=NSEW)

                        Label(hvframe,text='USP').grid(row=1, column=0,columnspan=1,sticky=NSEW)
                        Label(hvframe,text='AH 508 AH 513 AH 515').grid(row=1, column=1,columnspan=3,sticky=NSEW)
                        Label(hvframe, text='Awaiting PI:', bg="yellow").grid(row=2, column=0, columnspan=1, sticky=NSEW)
                        # uspinput = Entry(hvframe,textvariable=usptbc)
                        # uspinput.bind("<Return>", (lambda event:resultgen(usptbc.get(),4,2,0)))
                        uspinput = Label(hvframe, text='Please Refresh')
                        uspinput.grid(row=2, column=2,columnspan=1,sticky=NSEW)

                        Label(hvframe, text='DSP').grid(row=3, column=0,columnspan=1,sticky=NSEW)
                        Label(hvframe, text='AH 5 AH 5 AH 5 AH 51 AH 42 AH 5 AH 55 AH 51').grid(row=3, column=1,columnspan=3,sticky=NSEW)
                        Label(hvframe, text='Awaiting PI:', bg="yellow").grid(row=4, column=0,columnspan=1,sticky=NSEW)
                        # dspinput = Entry(hvframe,textvariable=dsptbc)
                        # dspinput.bind("<Return>", (lambda event: resultgen(dsptbc.get(), 4, 4, 0)))
                        dspinput = Label(hvframe, text='Please Refresh')
                        dspinput.grid(row=4, column=2,columnspan=1,sticky=NSEW)

                        Label(hvframe, text='SHF').grid(row=5, column=0,columnspan=1,sticky=NSEW)
                        Label(hvframe, text='AH 5 AH 4 AH 3').grid(row=5, column=1,columnspan=3,sticky=NSEW)
                        Label(hvframe, text='Awaiting PI:', bg="yellow").grid(row=6, column=0,columnspan=1,sticky=NSEW)
                        # shfinput = Entry(hvframe,textvariable=shftbc)
                        # shfinput.bind("<Return>", (lambda event: resultgen(shftbc.get(), 4, 6, 0)))
                        shfinput = Label(hvframe, text='Please Refresh')
                        shfinput.grid(row=6, column=2,columnspan=1,sticky=NSEW)

                        Label(hvframe, text='COMMON').grid(row=7, column=0,columnspan=1,sticky=NSEW)
                        Label(hvframe, text='AH 519 AH520').grid(row=7, column=1,columnspan=3,sticky=NSEW)
                        Label(hvframe, text='Awaiting PI:', bg="yellow").grid(row=8, column=0,columnspan=1,sticky=NSEW)
                        # cominput = Entry(hvframe,textvariable=comtbc)
                        # cominput.bind("<Return>", (lambda event: resultgen(comtbc.get(), 4, 8, 0)))
                        cominput = Label(hvframe, text='Please Refresh')
                        cominput.grid(row=8, column=2,columnspan=1,sticky=NSEW)

                        Label(hvframe, text='UTILITIES',relief="ridge",borderwidth = 1).grid(row=0, column=5, columnspan=6, sticky=NSEW)

                        Label(hvframe,text='WFI 1').grid(row=1, column=5,columnspan=1,sticky=NSEW)
                        Label(hvframe,text='STC11 MS&T MP BP').grid(row=1, column=6,columnspan=3,sticky=NSEW)
                        Label(hvframe,text='Awaiting PI:', bg="yellow").grid(row=2, column=5,columnspan=1,sticky=NSEW)
                        # wfi21input = Entry(hvframe, textvariable=wfi21tbc)
                        # # wfi21input.bind("<Return>", (lambda event: resultgen(wfi21tbc.get(), 20, 2, 5)))
                        wfi21input = Label(hvframe, text='Please Refresh')
                        wfi21input.grid(row=2, column=7, columnspan=1, sticky=NSEW)

                        Label(hvframe, text='WFI 2').grid(row=3, column=5,columnspan=1,sticky=NSEW)
                        Label(hvframe, text='CIP USP-DSP').grid(row=3, column=6,columnspan=3,sticky=NSEW)
                        Label(hvframe, text='Awaiting PI:', bg="yellow").grid(row=4, column=5,columnspan=1,sticky=NSEW)
                        # wfi22input = Entry(hvframe, textvariable=wfi22tbc)
                        # wfi22input.bind("<Return>", (lambda event: resultgen(wfi22tbc.get(), 20, 4, 5)))
                        wfi22input = Label(hvframe, text='Please Refresh')
                        wfi22input.grid(row=4, column=7, columnspan=1, sticky=NSEW)

                        Label(hvframe, text='WFI 3').grid(row=5, column=5,columnspan=1,sticky=NSEW)
                        Label(hvframe, text='STC1 STC3 BP2 DSP2').grid(row=5, column=6,columnspan=3,sticky=NSEW)
                        Label(hvframe, text='Awaiting PI:', bg="yellow").grid(row=6, column=5,columnspan=1,sticky=NSEW)
                        # wfi23input = Entry(hvframe, textvariable=wfi23tbc)
                        # wfi23input.bind("<Return>", (lambda event: resultgen(wfi23tbc.get(), 20, 6, 5)))
                        wfi23input = Label(hvframe, text='Please Refresh')
                        wfi23input.grid(row=6, column=7, columnspan=1, sticky=NSEW)

                        Label(hvframe, text='TMAC Tank').grid(row=7, column=5,columnspan=1,sticky=NSEW)
                        Label(hvframe, text='').grid(row=7, column=6,columnspan=3,sticky=NSEW)
                        Label(hvframe, text='Awaiting PI:', bg="yellow").grid(row=8, column=5,columnspan=1,sticky=NSEW)
                        # Label(hvframe, text='Input').grid(row=8, column=6,columnspan=3,sticky=NSEW)
                        # tmacinput = Entry(hvframe, textvariable=tmactbc)
                        # tmacinput.bind("<Return>", (lambda event: resultgen(tmactbc.get(), 20, 8, 5)))
                        tmacinput = Label(hvframe, text='Please Refresh')
                        tmacinput.grid(row=8, column=7, columnspan=1, sticky=NSEW)

                        Button(hvframe, text="Refresh\nPI DB", bg=btcolor,command=lambda:loadpidf(),relief=FLAT).grid(row=1, column=10, rowspan=8, sticky=NSEW)

                    elif dropselect.get() == 'Prod. Schedule':
                        def psaitem_selected(event):
                            def savebutton(batchcount,runcount,VBcount,Gcount,t20count,t50count,t500count,t1400count,t10000count,harvestcount,fillcount,t20Scount,t50Scount,t500Scount,t1400Scount,t10000Scount,harvestScount,fillScount,indexD):
                                psadf.at[indexD, 'Batch Count'] = str(batchcount.get("1.0", 'end-1c'))
                                psadf.at[indexD, 'Run'] = str(runcount.get("1.0", 'end-1c'))
                                psadf.at[indexD, 'G+-'] = str(Gcount.get())
                                psadf.at[indexD, 'Vial Break'] = str(VBcount.get())
                                psadf.at[indexD, '20 L'] = str(t20count.get("1.0", 'end-1c'))
                                psadf.at[indexD, '50 L'] = str(t50count.get("1.0", 'end-1c'))
                                psadf.at[indexD, '500 L'] = str(t500count.get("1.0", 'end-1c'))
                                psadf.at[indexD, '1.5 K'] = str(t1400count.get("1.0", 'end-1c'))
                                psadf.at[indexD, '10 K'] = str(t10000count.get("1.0", 'end-1c'))
                                psadf.at[indexD, 'Harvest'] = str(harvestcount.get("1.0", 'end-1c'))
                                psadf.at[indexD, 'Fill'] = str(fillcount.get("1.0", 'end-1c'))
                                psadf.at[indexD, '20 L S'] = str(t20Scount.get())
                                psadf.at[indexD, '50 L S'] = str(t50Scount.get())
                                psadf.at[indexD, '500 L S'] = str(t500Scount.get())
                                psadf.at[indexD, '1.5 K S'] = str(t1400Scount.get())
                                psadf.at[indexD, '10 K S'] = str(t10000Scount.get())
                                psadf.at[indexD, 'Harvest S'] = str(harvestScount.get())
                                psadf.at[indexD, 'Fill S'] = str(fillScount.get())
                                savepsadf()
                                loadmetricframe()
                                popframe.quit()
                                popframe.destroy()

                            def delbutton(indexD):
                                psadf.drop(indexD, inplace=True)
                                savepsadf()
                                loadmetricframe()
                                popframe.quit()
                                popframe.destroy()

                            def exitbutton():
                                popframe.quit()
                                popframe.destroy()

                            item_id = event.widget.focus()
                            item = event.widget.item(item_id)
                            values = item['values']
                            batchcountD = values[0]
                            runD = values[1]
                            twentyD = values[4]
                            fiftyD = values[5]
                            fivehundD = values[6]
                            onekfourD = values[7]
                            tenkD = values[8]
                            harvestD = values[9]
                            fillD = values[10]
                            indexD = values[18]

                            popframe = Toplevel(main_screen)
                            popframe.geometry("1080x240")
                            popframe.title('Batch: ' + str(runD))

                            Label(popframe,text = 'Batch Count:').place(relx=0.1, rely=0)
                            Label(popframe, text='Run Number:').place(relx=0.3, rely=0)
                            Label(popframe, text='Vial Break:').place(relx=0.5, rely=0)
                            Label(popframe, text='G±:').place(relx=0.7, rely=0)

                            batchcount = tk.Text(popframe)
                            batchcount.insert(INSERT, batchcountD)
                            batchcount.place(relx=0.1, rely=0.1,relheight=0.1,relwidth=0.1)

                            runcount = tk.Text(popframe)
                            runcount.insert(INSERT, runD)
                            runcount.place(relx=0.3, rely=0.1,relheight=0.1,relwidth=0.1)

                            VialbreakD = StringVar()
                            VialbreakD.set(values[3])
                            VBcount = OptionMenu(popframe, VialbreakD, 'Thaw', 'No Thaw')
                            VBcount.config(bg=bgcolor, relief=FLAT)
                            VBcount.place(relx=0.5, rely=0.1,relheight=0.1,relwidth=0.1)

                            gStatusD = StringVar()
                            gStatusD.set(values[2])
                            Gcount = OptionMenu(popframe, gStatusD, 'Yes', 'No')
                            Gcount.config(bg=bgcolor, relief=FLAT)
                            Gcount.place(relx=0.7, rely=0.1,relheight=0.1,relwidth=0.1)

                            Label(popframe, text='20L:').place(relx=0.1, rely=0.3)
                            Label(popframe, text='50L:').place(relx=0.2, rely=0.3)
                            Label(popframe, text='500L:').place(relx=0.3, rely=0.3)
                            Label(popframe, text='1400L:').place(relx=0.4, rely=0.3)
                            Label(popframe, text='10000L:').place(relx=0.5, rely=0.3)
                            Label(popframe, text='Harvest:').place(relx=0.6, rely=0.3)
                            Label(popframe, text='Fill:').place(relx=0.7, rely=0.3)

                            t20count = tk.Text(popframe)
                            t20count.insert(INSERT, twentyD)
                            t20count.place(relx=0.1, rely=0.4,relheight=0.1,relwidth=0.1)

                            t50count = tk.Text(popframe)
                            t50count.insert(INSERT, fiftyD)
                            t50count.place(relx=0.2, rely=0.4,relheight=0.1,relwidth=0.1)

                            t500count = tk.Text(popframe)
                            t500count.insert(INSERT, fivehundD)
                            t500count.place(relx=0.3, rely=0.4,relheight=0.1,relwidth=0.1)

                            t1400count = tk.Text(popframe)
                            t1400count.insert(INSERT, onekfourD)
                            t1400count.place(relx=0.4, rely=0.4,relheight=0.1,relwidth=0.1)

                            t10000count = tk.Text(popframe)
                            t10000count.insert(INSERT, tenkD)
                            t10000count.place(relx=0.5, rely=0.4,relheight=0.1,relwidth=0.1)

                            harvestcount = tk.Text(popframe)
                            harvestcount.insert(INSERT, harvestD)
                            harvestcount.place(relx=0.6, rely=0.4,relheight=0.1,relwidth=0.1)

                            fillcount = tk.Text(popframe)
                            fillcount.insert(INSERT, fillD)
                            fillcount.place(relx=0.7, rely=0.4,relheight=0.1,relwidth=0.1)

                            Label(popframe, text='On track?:').place(relx=0.1, rely=0.5)
                            Label(popframe, text='On track?:').place(relx=0.2, rely=0.5)
                            Label(popframe, text='On track?:').place(relx=0.3, rely=0.5)
                            Label(popframe, text='On track?:').place(relx=0.4, rely=0.5)
                            Label(popframe, text='On track?:').place(relx=0.5, rely=0.5)
                            Label(popframe, text='On track?:').place(relx=0.6, rely=0.5)
                            Label(popframe, text='On track?:').place(relx=0.7, rely=0.5)

                            twentySD = StringVar()
                            twentySD.set(values[11])
                            t20Scount = OptionMenu(popframe, twentySD, 'Y', 'N')
                            t20Scount.config(bg=bgcolor, relief=FLAT)
                            t20Scount.place(relx=0.1, rely=0.6,relheight=0.1,relwidth=0.1)

                            fiftySD = StringVar()
                            fiftySD.set(values[12])
                            t50Scount = OptionMenu(popframe, fiftySD, 'Y', 'N')
                            t50Scount.config(bg=bgcolor, relief=FLAT)
                            t50Scount.place(relx=0.2, rely=0.6,relheight=0.1,relwidth=0.1)

                            fivehundSD = StringVar()
                            fivehundSD.set(values[13])
                            t500Scount = OptionMenu(popframe, fivehundSD, 'Y', 'N')
                            t500Scount.config(bg=bgcolor, relief=FLAT)
                            t500Scount.place(relx=0.3, rely=0.6,relheight=0.1,relwidth=0.1)

                            onekfourSD = StringVar()
                            onekfourSD.set(values[14])
                            t1400Scount = OptionMenu(popframe, onekfourSD, 'Y', 'N')
                            t1400Scount.config(bg=bgcolor, relief=FLAT)
                            t1400Scount.place(relx=0.4, rely=0.6,relheight=0.1,relwidth=0.1)

                            tenkSD = StringVar()
                            tenkSD.set(values[15])
                            t10000Scount = OptionMenu(popframe, tenkSD, 'Y', 'N')
                            t10000Scount.config(bg=bgcolor, relief=FLAT)
                            t10000Scount.place(relx=0.5, rely=0.6,relheight=0.1,relwidth=0.1)

                            harvestSD = StringVar()
                            harvestSD.set(values[16])
                            harvestScount = OptionMenu(popframe, harvestSD, 'Y', 'N')
                            harvestScount.config(bg=bgcolor, relief=FLAT)
                            harvestScount.place(relx=0.6, rely=0.6,relheight=0.1,relwidth=0.1)

                            fillSD = StringVar()
                            fillSD.set(values[17])
                            fillScount = OptionMenu(popframe, fillSD, 'Y', 'N')
                            fillScount.config(bg=bgcolor, relief=FLAT)
                            fillScount.place(relx=0.7, rely=0.6,relheight=0.1,relwidth=0.1)


                            Button(popframe, text="Save", command=lambda: savebutton(batchcount,runcount,VialbreakD,gStatusD,t20count,t50count,t500count,t1400count,t10000count,harvestcount,fillcount,twentySD,fiftySD,fivehundSD,onekfourSD,tenkSD,harvestSD,fillSD,indexD), relief=FLAT).place(rely=0.8,relx=0.0,relheight=buttonrelheight * 2,relwidth=.3)
                            Button(popframe, text="Delete", command=lambda: delbutton(indexD), relief=FLAT).place(
                                relx=0.5 - 0.15, rely=0.8, relheight=buttonrelheight * 2, relwidth=.3)
                            Button(popframe, text="Close", command=exitbutton, relief=FLAT).place(relx=0.7, rely=0.8,
                                                                                                  relheight=buttonrelheight * 2,
                                                                                                  relwidth=.3)

                            popframe.mainloop()

                        def new_psa_entry_button():

                            def savebutton(batchcount, runcount, VBcount, Gcount, t20count, t50count, t500count,
                                           t1400count, t10000count, harvestcount, fillcount, t20Scount, t50Scount,
                                           t500Scount, t1400Scount, t10000Scount, harvestScount, fillScount):
                                loadpsadf()
                                indexD = psadf.index.max() + 1
                                psadf.at[indexD, 'Batch Count'] = str(batchcount.get("1.0", 'end-1c'))
                                psadf.at[indexD, 'Run'] = str(runcount.get("1.0", 'end-1c'))
                                psadf.at[indexD, 'G+-'] = str(Gcount.get())
                                psadf.at[indexD, 'Vial Break'] = str(VBcount.get())
                                psadf.at[indexD, '20 L'] = str(t20count.get("1.0", 'end-1c'))
                                psadf.at[indexD, '50 L'] = str(t50count.get("1.0", 'end-1c'))
                                psadf.at[indexD, '500 L'] = str(t500count.get("1.0", 'end-1c'))
                                psadf.at[indexD, '1.5 K'] = str(t1400count.get("1.0", 'end-1c'))
                                psadf.at[indexD, '10 K'] = str(t10000count.get("1.0", 'end-1c'))
                                psadf.at[indexD, 'Harvest'] = str(harvestcount.get("1.0", 'end-1c'))
                                psadf.at[indexD, 'Fill'] = str(fillcount.get("1.0", 'end-1c'))
                                psadf.at[indexD, '20 L S'] = str(t20Scount.get())
                                psadf.at[indexD, '50 L S'] = str(t50Scount.get())
                                psadf.at[indexD, '500 L S'] = str(t500Scount.get())
                                psadf.at[indexD, '1.5 K S'] = str(t1400Scount.get())
                                psadf.at[indexD, '10 K S'] = str(t10000Scount.get())
                                psadf.at[indexD, 'Harvest S'] = str(harvestScount.get())
                                psadf.at[indexD, 'Fill S'] = str(fillScount.get())
                                savepsadf()
                                loadmetricframe()
                                popframe.quit()
                                popframe.destroy()

                            def exitbutton():
                                popframe.quit()
                                popframe.destroy()

                            popframe = Toplevel(main_screen)
                            popframe.geometry("1080x240")
                            popframe.title('New Batch')

                            Label(popframe, text='Batch Count:').place(relx=0.1, rely=0)
                            Label(popframe, text='Run Number:').place(relx=0.3, rely=0)
                            Label(popframe, text='Vial Break:').place(relx=0.5, rely=0)
                            Label(popframe, text='G±:').place(relx=0.7, rely=0)

                            batchcount = tk.Text(popframe)
                            batchcount.insert(INSERT, '')
                            batchcount.place(relx=0.1, rely=0.1, relheight=0.1, relwidth=0.1)

                            runcount = tk.Text(popframe)
                            runcount.insert(INSERT, '')
                            runcount.place(relx=0.3, rely=0.1, relheight=0.1, relwidth=0.1)

                            VialbreakD = StringVar()
                            # VialbreakD.set(values[3])
                            VBcount = OptionMenu(popframe, VialbreakD, 'Thaw', 'No Thaw')
                            VBcount.config(bg=bgcolor, relief=FLAT)
                            VBcount.place(relx=0.5, rely=0.1, relheight=0.1, relwidth=0.1)

                            gStatusD = StringVar()
                            # gStatusD.set('Yes')
                            Gcount = OptionMenu(popframe, gStatusD, 'Yes', 'No')
                            Gcount.config(bg=bgcolor, relief=FLAT)
                            Gcount.place(relx=0.7, rely=0.1, relheight=0.1, relwidth=0.1)

                            Label(popframe, text='20L:').place(relx=0.1, rely=0.3)
                            Label(popframe, text='50L:').place(relx=0.2, rely=0.3)
                            Label(popframe, text='500L:').place(relx=0.3, rely=0.3)
                            Label(popframe, text='1400L:').place(relx=0.4, rely=0.3)
                            Label(popframe, text='10000L:').place(relx=0.5, rely=0.3)
                            Label(popframe, text='Harvest:').place(relx=0.6, rely=0.3)
                            Label(popframe, text='Fill:').place(relx=0.7, rely=0.3)

                            t20count = tk.Text(popframe)
                            t20count.insert(INSERT, '')
                            t20count.place(relx=0.1, rely=0.4, relheight=0.1, relwidth=0.1)

                            t50count = tk.Text(popframe)
                            t50count.insert(INSERT, '')
                            t50count.place(relx=0.2, rely=0.4, relheight=0.1, relwidth=0.1)

                            t500count = tk.Text(popframe)
                            t500count.insert(INSERT, '')
                            t500count.place(relx=0.3, rely=0.4, relheight=0.1, relwidth=0.1)

                            t1400count = tk.Text(popframe)
                            t1400count.insert(INSERT, '')
                            t1400count.place(relx=0.4, rely=0.4, relheight=0.1, relwidth=0.1)

                            t10000count = tk.Text(popframe)
                            t10000count.insert(INSERT, '')
                            t10000count.place(relx=0.5, rely=0.4, relheight=0.1, relwidth=0.1)

                            harvestcount = tk.Text(popframe)
                            harvestcount.insert(INSERT, '')
                            harvestcount.place(relx=0.6, rely=0.4, relheight=0.1, relwidth=0.1)

                            fillcount = tk.Text(popframe)
                            fillcount.insert(INSERT, '')
                            fillcount.place(relx=0.7, rely=0.4, relheight=0.1, relwidth=0.1)

                            Label(popframe, text='On track?:').place(relx=0.1, rely=0.5)
                            Label(popframe, text='On track?:').place(relx=0.2, rely=0.5)
                            Label(popframe, text='On track?:').place(relx=0.3, rely=0.5)
                            Label(popframe, text='On track?:').place(relx=0.4, rely=0.5)
                            Label(popframe, text='On track?:').place(relx=0.5, rely=0.5)
                            Label(popframe, text='On track?:').place(relx=0.6, rely=0.5)
                            Label(popframe, text='On track?:').place(relx=0.7, rely=0.5)

                            twentySD = StringVar()
                            twentySD.set('Y')
                            t20Scount = OptionMenu(popframe, twentySD, 'Y', 'N')
                            t20Scount.config(bg=bgcolor, relief=FLAT)
                            t20Scount.place(relx=0.1, rely=0.6, relheight=0.1, relwidth=0.1)

                            fiftySD = StringVar()
                            fiftySD.set('Y')
                            t50Scount = OptionMenu(popframe, fiftySD, 'Y', 'N')
                            t50Scount.config(bg=bgcolor, relief=FLAT)
                            t50Scount.place(relx=0.2, rely=0.6, relheight=0.1, relwidth=0.1)

                            fivehundSD = StringVar()
                            fivehundSD.set('Y')
                            t500Scount = OptionMenu(popframe, fivehundSD, 'Y', 'N')
                            t500Scount.config(bg=bgcolor, relief=FLAT)
                            t500Scount.place(relx=0.3, rely=0.6, relheight=0.1, relwidth=0.1)

                            onekfourSD = StringVar()
                            onekfourSD.set('Y')
                            t1400Scount = OptionMenu(popframe, onekfourSD, 'Y', 'N')
                            t1400Scount.config(bg=bgcolor, relief=FLAT)
                            t1400Scount.place(relx=0.4, rely=0.6, relheight=0.1, relwidth=0.1)

                            tenkSD = StringVar()
                            tenkSD.set('Y')
                            t10000Scount = OptionMenu(popframe, tenkSD, 'Y', 'N')
                            t10000Scount.config(bg=bgcolor, relief=FLAT)
                            t10000Scount.place(relx=0.5, rely=0.6, relheight=0.1, relwidth=0.1)

                            harvestSD = StringVar()
                            harvestSD.set('Y')
                            harvestScount = OptionMenu(popframe, harvestSD, 'Y', 'N')
                            harvestScount.config(bg=bgcolor, relief=FLAT)
                            harvestScount.place(relx=0.6, rely=0.6, relheight=0.1, relwidth=0.1)

                            fillSD = StringVar()
                            fillSD.set('Y')
                            fillScount = OptionMenu(popframe, fillSD, 'Y', 'N')
                            fillScount.config(bg=bgcolor, relief=FLAT)
                            fillScount.place(relx=0.7, rely=0.6, relheight=0.1, relwidth=0.1)

                            Button(popframe, text="Save",
                                   command=lambda: savebutton(batchcount, runcount, VialbreakD, gStatusD, t20count,
                                                              t50count, t500count, t1400count, t10000count,
                                                              harvestcount, fillcount, twentySD, fiftySD, fivehundSD,
                                                              onekfourSD, tenkSD, harvestSD, fillSD),
                                   relief=FLAT).place(rely=0.8, relx=0.0, relheight=buttonrelheight * 2, relwidth=.3)

                            Button(popframe, text="Close", command=exitbutton, relief=FLAT).place(relx=0.7, rely=0.8,
                                                                                                  relheight=buttonrelheight * 2,
                                                                                                  relwidth=.3)

                            popframe.mainloop()

                        def fixed_map(option):
                            # Returns the style map for 'option' with any styles starting with
                            # ("!disabled", "!selected", ...) filtered out

                            # style.map() returns an empty list for missing options, so this should
                            # be future-safe
                            return [elm for elm in style.map("Treeview", query_opt=option)
                                    if elm[:2] != ("!disabled", "!selected")]

                        style = ttk.Style()
                        style.map("Treeview",
                                  foreground=fixed_map("foreground"),
                                  background=fixed_map("background"))

                        loadpsadf()

                        hvframe = LabelFrame(metricframe, text='Prod Schedule')
                        hvframe.place(relx=0, rely=0.07, relheight=.9, relwidth=1)
                        # Label(metricframe).place(relx=0, rely=0)
                        tableframe = LabelFrame(hvframe)
                        tableframe.place(relx=0, rely=0, relheight=1, relwidth=.89)

                        columns = ('#1', '#2', '#3', '#4', '#5', '#6', '#7','#8','#9','#10','#11','#12','#13','#14','#15','#16','#17','#18')
                        treeps = ttk.Treeview(tableframe, columns=columns, show='headings')
                        treeps.heading('#1', text='Batch Count')
                        treeps.column('#1', width=50)
                        treeps.heading('#2', text='Run')
                        treeps.column('#2', width=50)
                        treeps.heading('#3', text='G±')
                        treeps.column('#3', width=50)
                        treeps.heading('#4', text='Vial Break')
                        treeps.column('#4', width=50)
                        treeps.heading('#5', text='20L')
                        treeps.column('#5', width=50)
                        treeps.heading('#6', text='50L')
                        treeps.column('#6', width=50)
                        treeps.heading('#7', text='500L')
                        treeps.column('#7', width=50)
                        treeps.heading('#8', text='1400L')
                        treeps.column('#8', width=50)
                        treeps.heading('#9', text='10000L')
                        treeps.column('#9', width=50)
                        treeps.heading('#10', text='Harvest')
                        treeps.column('#10', width=50)
                        treeps.heading('#11', text='Fill')
                        treeps.column('#11', width=50)
                        treeps.heading('#12', text='20L S')
                        treeps.column('#12', stretch=NO, minwidth=0, width=0)
                        treeps.heading('#13', text='50L S')
                        treeps.column('#13', stretch=NO, minwidth=0, width=0)
                        treeps.heading('#14', text='500L S')
                        treeps.column('#14', stretch=NO, minwidth=0, width=0)
                        treeps.heading('#15', text='1400L S')
                        treeps.column('#15', stretch=NO, minwidth=0, width=0)
                        treeps.heading('#16', text='10000L S')
                        treeps.column('#16', stretch=NO, minwidth=0, width=0)
                        treeps.heading('#17', text='Harvest S')
                        treeps.column('#17', stretch=NO, minwidth=0, width=0)
                        treeps.heading('#18', text='Fill S')
                        treeps.column('#18', stretch=NO, minwidth=0, width=0)

                        treeps.place(relx=0, rely=0, relheight=1, relwidth=1)

                        for index, row in psadf.iterrows():
                            if 'N' in list(row)[11:18]:
                                treeps.insert('', 0, text=index, values=list(row), tags = ('N'))

                            else:
                                treeps.insert('', 0, text=index, values=list(row), tags=('Y'))

                        treeps.tag_configure('Y', foreground='green')
                        treeps.tag_configure('N', foreground='red')

                        # scrollbar
                        scrollbary = ttk.Scrollbar(hvframe, orient=tk.VERTICAL, command=treeps.yview)
                        treeps.configure(yscroll=scrollbary.set)
                        scrollbary.place(relx=0.89, rely=0, relheight=1, relwidth=.03)

                        scrollbarx = ttk.Scrollbar(hvframe, orient=tk.HORIZONTAL, command=treeps.xview)
                        treeps.configure(xscroll=scrollbarx.set)
                        scrollbarx.place(relx=0, rely=.95, relheight=.05, relwidth=.89)

                        Button(hvframe, text="Add new \nentry", command=lambda: new_psa_entry_button(), bg=btcolor,
                               relief=FLAT).place(
                            relx=1 - 0.075, rely=0, relheight=1, relwidth=.07)

                        treeps.bind('<Double-Button-1>', psaitem_selected)

                dropselect = StringVar(metricframe)
                dropselect.set('HVAC and Utilities')
                options_list = ['HVAC and Utilities','Prod. Schedule','PM/CM Status','Batch Release Status','Shipping Status']

                Label(metricframe, text='Selection:').place(relx=0, rely=0)
                pageselect = OptionMenu(metricframe, dropselect, *options_list,
                                        command=lambda _: loadmetricframe())
                pageselect.config(bg=bgcolor, relief=FLAT)
                pageselect.place(relx=0.07, rely=0.0, relheight=0.07)
                loadmetricframe()

            metricframe = LabelFrame(labelframe)
            metricframe.place(relx=0.08, rely=0.405, relheight=0.585, relwidth=.84 + .075)

            if subgroup.get() == 'People' and dropselect.get() == '2':
                tier2people()
            elif subgroup.get() == 'Safety' and dropselect.get() == '2':
                tier2safety()
            elif subgroup.get() == 'Quality' and dropselect.get() == '2':
                tier2quality()
            elif subgroup.get() == 'Delivery' and dropselect.get() == '2':
                tier2delivery()
            elif subgroup.get() == 'Cost' and dropselect.get() == '2':
                tier2cost()

    loaddf()

    main_screen.title("Tier : " + dropselect.get())
    labelframe = LabelFrame(main_screen, text=subgroup.get(), bg=bgcolor)
    labelframe.place(relx=offset, rely=0.05 + buttonrelheight, relheight=1 - 0.2 , relwidth=1 - offset * 2)

    drawdayframe()
    drawtreeview()
    drawmetrics()

def loadmain():
    global main_screen
    global labelframe
    global offset
    global buttonrelwidth
    global buttonrelheight
    global date
    global dropselect
    global subgroup
    global tdate
    global bgcolor
    global btcolor
    global fgcolor


    btcolor = 'gray90'
    bgcolor = '#FAFAFA'
    fgcolor = '#000000'

    date = dt.datetime.now()
    tdate = str(date.strftime("%d"))+' '+str(date.strftime("%b"))+' '+str(date.strftime("%y"))
    main_screen = Tk()
    main_screen.geometry("1280x720")
    main_screen.title("Tier Welcome")
    main_screen.configure(bg=bgcolor)
    defaultFont = font.nametofont("TkDefaultFont")
    defaultFont.configure(family="Calibri", size=11, weight=font.BOLD)
    buttonrelwidth = 0.1
    buttonrelheight = 0.05
    offset = 0.05
    dropselect = StringVar(main_screen)
    subgroup = StringVar(main_screen)
    dropselect.set('2')
    subgroup.set('People')


    Label(main_screen, text='Tier selection: ', bg=bgcolor).place(relx=.8, rely=0.01)
    options_list = ['USP SHO','DSP1 SHO','DSP2 SHO','SHF SHO','USP QAT','DSP1 QAT','DSP2 QAT','SHF QAT', '2', '3']
    tierselect = OptionMenu(main_screen, dropselect, *options_list, command=lambda _: [loadframe(),drawbuttons()])
    tierselect.config(bg=bgcolor, relief=FLAT)
    tierselect.place(relx=.9, rely=0.01, relheight=0.03)

    Label(main_screen, text='Date: '+tdate, bg=bgcolor).place(relx=offset, rely=0.01)

    # why is this logo canvas so hard to get working lol
    # canvas = Canvas(width=200, height=32,)
    # canvas.create_image(200, 32, image=logofilename, anchor='nw')
    # canvas.place(relx=.1, rely=0.95)

    loadframe()
    drawbuttons()

    main_screen.mainloop()

if __name__ == '__main__':
    loadmain()