import pandas as pd
df = pd.read_excel("https://docs.google.com/spreadsheets/d/e/2PACX-1vSAH0q4IKYBmA46IT2IbrjpuxwwmOlKUcAZSaaFWa_iiv0Iv6sER08JDWAy0wJ8mjdmobE9OORIyjft/pub?output=xlsx")

eventindex1 = 4 #First index in the spreadsheet where event abbreviations begin
eventindex2 = 12#Last index in the spreadsheet where event abbreviations begin

def checkdupeevents(df):
    good = True
    for i in range(len(df.index)):
        n = 0
        for j in range(eventindex1,eventindex2+1):
            if isinstance(df.iat[i,j],str):
                n = n + 1
        if n == 0:
            print("No events detected on row " + str(i + 2))
            good = False
        elif n > 1:
            print("Too many events detected on row " + str(i + 2))
            good = False
    if good:
        print("No problems were detected!")


checkdupeevents(df)
