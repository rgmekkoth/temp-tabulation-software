import pandas as pd

df = pd.read_excel("https://docs.google.com/spreadsheets/d/e/2PACX-1vQ5IDq9qak5JmrcXra2_UKYR3HRegeWccDSnLEtgsc4cvjXarEtZ3uVrWC4fPXf_A/pub?output=xlsx")

#returns exam score for a student given a student id.
def getexam(studentid):
    for i in range(len(df.index)):
        if (df.iat[i,1]) == studentid:
            return (100*df.iat[i,4])
    return 0

print(getexam(16885))
print(getexam(20111))
print(getexam(39321))
