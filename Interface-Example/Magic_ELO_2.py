# -*- coding: utf-8 -*-
"""
Created on Fri Feb 16 20:58:15 2024

@author: Alexa Nemeth

numpy version:  1.20.3
pandas version:  1.4.3
tkinter version: 8.6.11

"""
#inport packages
import numpy as np
import pandas as pd
import tkinter as tk
import tkinter.filedialog
from tkinter import *
from functools import partial
from  tkinter import ttk


#create function to call with button to open elo match history file
def open_file():
    #read the file in
    filepath = tk.filedialog.askopenfilename()  
    #call function to create radio buttons for decks and rest of user interface / calculations 
    create_buttons(filepath)
    #remove button once filepath is read
    import_xlsx.destroy()
    return

#Function to create deck radio buttons and submit results for calculation    
def create_buttons(filepath):
    #create variable for radio button inputs
    deck1 = StringVar()
    deck1_result = StringVar()
    deck2 = StringVar()
    deck2_result = StringVar()
    #read in the elo and hitsory worksheets
    elo = pd.read_excel(filepath, sheet_name='Ranking')
    history = pd.read_excel(filepath, sheet_name='Match History')
    #create list with all deck names
    decks = elo.columns.values.tolist()
    #create buttons for first deck
    for r in range(elo.shape[1]):
        tk.Radiobutton(canvas,text= decks[r],value = decks[r],variable = deck1,tristatevalue=0, command=None).grid(column=1, row=r+2, pady=2, padx=2)
    #create buttons for result of the first deck
    tk.Radiobutton(canvas, text = "Win", value ="1",variable = deck1_result, tristatevalue=1,command=None).grid(column=2, row=2, pady=2, padx=2)
    tk.Radiobutton(canvas, text = "Loss", value ="0",variable = deck1_result,tristatevalue=2,command=None).grid(column=2, row=3, pady=2, padx=2)
    tk.Radiobutton(canvas, text = "Tie", value ="0.5",variable = deck1_result,tristatevalue=3,command=None).grid(column=2, row=4, pady=2, padx=2)
    #Create buttons for second deck
    for r in range(elo.shape[1]):
        tk.Radiobutton(canvas, text= decks[r],value = decks[r],variable = deck2,tristatevalue=4,command=None).grid(column=3, row=r+2, pady=2, padx=2)
    #create buttoons for result of second deck
    tk.Radiobutton(canvas, text = "Win", value ="1",variable = deck2_result,tristatevalue=5,command=None).grid(column=4, row=2, pady=2, padx=2)
    tk.Radiobutton(canvas, text = "Loss", value ="0",variable = deck2_result,tristatevalue=6,command=None).grid(column=4, row=3, pady=2, padx=2)
    tk.Radiobutton(canvas, text = "Tie", value ="0.5",variable = deck2_result,tristatevalue=7,command=None).grid(column=4, row=4, pady=2, padx=2)
    
    #create button hover color change
    def resultButton_hover(x):
        resultButton['background'] = '#d6eaf8'
    def resultButton_not_hover(x):
        resultButton['background'] = '#ebf5fb'
    
    #create button to report selected radio buttons and start ELO calculation
    resultButton = tk.Button(canvas, text="Report Results", command=partial(elo_calculation,filepath, elo, history,deck1,deck1_result,deck2,deck2_result),relief = GROOVE)
    resultButton.grid(column=1, row = 1)
    resultButton.bind("<Enter>", resultButton_hover)
    resultButton.bind("<Leave>", resultButton_not_hover)
    return

#Calculate the new ELO    
def elo_calculation(filepath, elo, history,deck1,deck1_result,deck2,deck2_result):
    #Set fixed variables
    S = 400
    K = 32

    #create the expected outcome of match between Ra and Rb
    def E_pred(Ra, Rb):
        expected_winner = 1/(1+10**((Rb-Ra)/S))
        return expected_winner

    #Give new ranking to Ra or Rb after match
    def New_ranking(ranking, predict_outcome, O): 
        R_prime = ranking + K*(O-predict_outcome)
        return R_prime

    #First deck of match
    deck_name_1 = deck1.get()
    deck_1_O= deck1_result.get()
    #Second deck of match
    deck_name_2 = deck2.get()
    deck_2_O= deck2_result.get()

    #get the last rankings index 
    last_ranking_index = len(elo.index) -1

    #convert input to float num
    deck_1_O = float(deck_1_O)
    deck_2_O = float(deck_2_O)

    #get Current Rankings
    deck1_R = elo.iloc[last_ranking_index][deck_name_1]
    deck2_R = elo.iloc[last_ranking_index][deck_name_2]

    #calculate the predicted winners 
    deck_1_outcome_predict = E_pred(deck1_R, deck2_R)
    deck_2_outcome_predict = E_pred(deck2_R, deck1_R)

    #Calculate the decks new rankings
    deck_1_R_prime = round(New_ranking(deck1_R, deck_1_outcome_predict, deck_1_O),2)
    deck_2_R_prime = round(New_ranking(deck2_R, deck_2_outcome_predict, deck_2_O),2)

    #save the data to history
    elo2 = pd.DataFrame(elo.loc[last_ranking_index,:]).T
    elo = pd.concat([elo,elo2], ignore_index = True)
    elo.at[last_ranking_index+1, deck_name_1] = deck_1_R_prime
    elo.at[last_ranking_index+1, deck_name_2] = deck_2_R_prime

    blank = pd.DataFrame(2, index=np.arange(1), columns=history.columns)
    history = pd.concat([history,blank], ignore_index = True)
    history.at[last_ranking_index+1, deck_name_1] = deck_1_O
    history.at[last_ranking_index+1, deck_name_2] = deck_2_O

    #Export elo to spreadsheet
    with pd.ExcelWriter(filepath) as writer:
        elo.to_excel(writer, sheet_name="Ranking",index=False)
        history.to_excel(writer, sheet_name="Match History",index=False)  
        
    #Call functions to display tables in user interface
    display_rankings(elo,history,filepath)
    display_elo(elo)
    return

#Function to display the ELO chart in user interface
def display_elo(elo):
    #Find the last ELO calculated and the deck names
    elolast = elo.iloc[-1]
    last_ELO = elolast.T
    last_ELO = pd.DataFrame(elolast)
    last_ELO.columns =['ELO']
    #Sort by highest to lowest ELO value
    last_ELO.sort_values(['ELO'], ascending=[False], inplace=True)
    
    #Create frame for Treeview
    elo_chart_frame = Frame(canvas)
    elo_chart_frame.grid(column = 5, row = 4, rowspan=9)
    
    #Create Tree view and give column labels and headings
    elo_chart = ttk.Treeview(elo_chart_frame)
    elo_chart['columns'] = ('Deck', 'ELO')
    
    elo_chart.column("#0", width=0,  stretch=NO)
    elo_chart.column("Deck",anchor=CENTER, width=80)
    elo_chart.column("ELO",anchor=CENTER,width=80)

    elo_chart.heading("#0",text="",anchor=CENTER)
    elo_chart.heading("Deck",text="Deck",anchor=CENTER)
    elo_chart.heading("ELO",text="ELO",anchor=CENTER)
    
    #Create row labels
    row_labels = last_ELO.index.values
    #read in data to the chart from the created dataframe
    for i in range(len(last_ELO)):
        elo_chart.insert(parent='',index='end',iid=i,text='',values=(row_labels[i],last_ELO.iloc[i,0]))
    #Set cart to grid    
    elo_chart.grid()    
    return

#Create chart for Wins and Losses 
def display_rankings(elo,history,filepath):
    
    #create deck name labels
    elolast = elo.iloc[-1]
    wins = 0
    loss = 0
    record = elolast.T
    record = pd.DataFrame(record)
    record.columns =['Wins']
    record['Wins'] = 0
    record['Losses'] = 0
    
    #total up wins and losses for each deck based on historical data
    for i in range(len(history.columns)):
        wins = 0
        loss = 0
        for j in range(len(history)):
            history.iloc()[j,i]
            if history.iloc[j,i] == 1:
                wins +=1
            elif history.iloc[j,i] == 0:
                loss +=1
        record.iloc[i,0] = wins
        record.iloc[i,1] = loss
    
    #sort by most wins and least losses
    record.sort_values(['Wins', 'Losses'], ascending=[False, True], inplace=True)
    #Create Tree view and give column labels and headings
    history_chart_frame = Frame(canvas)
    history_chart_frame.grid(column = 5, row = 12, rowspan=9)
    history_chart = ttk.Treeview(history_chart_frame)
    history_chart['columns'] = ('Deck', 'Wins', 'Losses')

    history_chart.column("#0", width=0,  stretch=NO)
    history_chart.column("Deck",anchor=CENTER, width=80)
    history_chart.column("Wins",anchor=CENTER,width=80)
    history_chart.column("Losses",anchor=CENTER,width=80)

    history_chart.heading("#0",text="",anchor=CENTER)
    history_chart.heading("Deck",text="Deck",anchor=CENTER)
    history_chart.heading("Wins",text="Wins",anchor=CENTER)
    history_chart.heading("Losses",text="Losses",anchor=CENTER)
    #Create row labels
    row_labels = record.index.values
    #read in data to the chart from the created dataframe
    for i in range(len(record)):
        history_chart.insert(parent='',index='end',iid=i,text='',values=(row_labels[i],record.iloc[i,0],record.iloc[i,1]))
    #Set to grid
    history_chart.grid()
    
    #Create hover color change for button
    def reset_hover(x):
        reset_reporting['background'] = '#ebf5fb'
    def reset_not_hover(x):
        reset_reporting['background'] = '#d6eaf8'
    #Create button to reset data and report another match    
    reset_reporting = Button(canvas, text='Clear Input and Report Next Match', command=partial(resetAll,filepath),relief = GROOVE)
    reset_reporting.grid(column=5, row = 1)
    reset_reporting.bind("<Enter>", reset_hover)
    reset_reporting.bind("<Leave>", reset_not_hover)
    return

#Function to reset all but filepath data
def resetAll(filepath):
    #delete inputs
    canvas.delete('all')
    #re-use same filepath and call button function again
    create_buttons(filepath)
    return

#create button hover color change
def import_hover(x):
    import_xlsx['background'] = '#ebf5fb'
def import_not_hover(x):
    import_xlsx['background'] = '#d6eaf8'
    
#create tkinter window and canvas to reset data to report multiple matches
mtg_elo_window = tk.Tk()
canvas = Canvas(mtg_elo_window)
canvas.pack()
mtg_elo_window.geometry('750x800+0+35')
mtg_elo_window.title('Magic ELO')

#Welcome mesage and button to start off importing historical data file
welcome_message = tk.Label(canvas, text="Welcome! Please import xlsx elo history file to get started")
welcome_message.grid(row=0, column=1, columnspan = 10)
import_xlsx = tk.Button(canvas, relief = GROOVE, height=1, width=10, text="Import Data", command=open_file)
import_xlsx.grid(row=1, column=1)
import_xlsx.bind("<Enter>", import_hover)
import_xlsx.bind("<Leave>", import_not_hover)
#Start main window
mtg_elo_window.mainloop()