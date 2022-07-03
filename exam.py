import pandas as pd

examdf = pd.read_excel("https://docs.google.com/spreadsheets/d/e/2PACX-1vQ5IDq9qak5JmrcXra2_UKYR3HRegeWccDSnLEtgsc4cvjXarEtZ3uVrWC4fPXf_A/pub?output=xlsx")

#returns exam score for a student given a student id. 
def getexam(studentid):
    for i in range(len(examdf.index)):
        if (examdf.iat[i,1]) == studentid:
            return round(100*((examdf.iat[i,5])/45),2)
    return 0

print(getexam(16885))
print(getexam(20111))
print(getexam(39321))
