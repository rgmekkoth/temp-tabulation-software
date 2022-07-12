import pandas as pd
#identify each unique google sheets files, change this to access a new sheet
sheet_id = '1P1yuimIAYmnfko1O89W51BUdeTszjj-KJwjIk8q2txM'

df = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv")

events = ["ACT", "AAM", "ASM", "BFS", "BSM", "ENT", "FMS", "HLM", "HRM", "MCS", "QSRM", "RFSM", "RMS", "SEM"]
participant = []
class ISR:
    #event typers: ACT, AAM, ASM, BFS, BSM, ENT, FMS, HLM, HRM, MCS, QSRM, RFSM, RMS, SEM
    def __init__(self,exam, r1, r2, eventType, id):
        self.exam = exam
        self.r1 = r1
        self.r2 = r2
        self.eventType = eventType
        self.id = id

    def calcISR(totalScore, position):
        test = df.iloc[position].tolist()
        exam = exam/100 * totalScore
        r1 = test[22]/100 * totalScore
        r2 = test[23]/100 * totalScore
        finalScore = exam+r1+r2
        
        return finalScore

    #get Participants along with their events
    def getParticipants(df):
        totalPart = len(df.index)-1
        participant = {}
        for i in range(totalPart):
            test = df.iloc[i].tolist()
            if test[12] == "":
                break
            id = test[4]
            if participant.indexOf(id)== -1:
                if not (events.indexOf(eventType) ==-1):
                    eventType = test[12]
                else:
                    eventType = "FALSE"
                participant[id] = eventType
        return participant

