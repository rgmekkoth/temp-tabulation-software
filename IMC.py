import pandas as pd
#identify each unique google sheets files, change this to access a new sheet
judgeSheet_id = '1rowe4EORKvSdPkUxhUAAihWttNgfjfBKOdH6lLidfP8'
studentSheet_id = '1iinNp8thmLsJQiTO3kYL48McR_uF1ClFkGZsgM0CCiQ'
studentDf = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{studentSheet_id}/export?format=csv")
judgeDf = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{judgeSheet_id}/export?format=csv")

events = ["IMCE", "IMCP", "IMCS"]
participant = []
eventRow = 11
teamNum = 2
writtenRow = 14
presRow = 15

#include exam later

class IMC:

    exam1 = 0
    exam2 = 0
    exam3 = 0
    present = None
    written = None
    
    def __init__(self, teamNum):
        self.teamNum = teamNum

    def findScore():
        global writtenRow
        global presRow
        global present, written, events


        row = judgeDf[judgeDf.isin([teamNum]).any(axis=1)]
        row = row.tolist()
        
        present = row[presRow]
        written = row[writtenRow]
    #calculate IMC final score. I tried using global variable to simplify the process but it has not worked 100% yet. the equation does work 
    def calcIMC():
        global exam1, exam2, exam3, present, written
        avgExam = (exam1+exam2+exam3)/3
        presentScore =present/100
        writtenScore = written/100
        finalScore = (presentScore*.35 + writtenScore*.35 + avgExam/100*.3)*100 #show final score as a percentage 
        return finalScore  
        
