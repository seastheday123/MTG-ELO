# -*- coding: utf-8 -*-
"""
Created on Sun Aug 28 11:33:57 2022

@author: Alexa

Automate ELO calculations for two player MTG games. Takes results of game and calculates new rankings. Exports match history and rankings to xlsx file. 

"""
#import packages
import numpy as np
import pandas as pd

#import historic rankings from elo spreadsheet
elo = pd.read_excel("ELO.xlsx")

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

#Take user input for new rankings
#First deck of match
deck_name_1 = input("Please enter name of deck 1:")
deck_1_O= input('Please enter outcome of game for the deck 1 [Win- 1; Draw- 0.5; Loss- 0]:')
#Second deck of match
deck_name_2 = input("Please enter name of deck 2:")
deck_2_O= input('Please enter outcome of game for the deck 2 [Win- 1; Draw- 0.5; Loss- 0]:')

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
deck_1_R_prime = New_ranking(deck1_R, deck_1_outcome_predict, deck_1_O)
deck_2_R_prime = New_ranking(deck2_R, deck_2_outcome_predict, deck_2_O)

#save the data to history
elo = elo.append(elo.loc[last_ranking_index], ignore_index=True)
elo.iloc[last_ranking_index+1][deck_name_1] = deck_1_R_prime
elo.iloc[last_ranking_index+1][deck_name_2] = deck_2_R_prime

#Export elo to spreadsheet
elo.to_excel('elo.xlsx',index=False)

#show the updated standings
print(elo)
print('New Rankings:')
print(deck_name_1," : ", deck_1_R_prime)
print(deck_name_2," : ", deck_2_R_prime)

