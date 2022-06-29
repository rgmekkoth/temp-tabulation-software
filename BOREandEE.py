
import pandas as pd
df = pd.read_excel("https://docs.google.com/spreadsheets/d/e/2PACX-1vSAH0q4IKYBmA46IT2IbrjpuxwwmOlKUcAZSaaFWa_iiv0Iv6sER08JDWAy0wJ8mjdmobE9OORIyjft/pub?output=xlsx")
class BOREGroup():
    def __init__(self, member1, oral, written, groupnumber, eventtype):
        self.members = [member1]
        self.groupnumber = groupnumber
        self.written = written
        self.oral = oral
        self.eventtype = eventtype

def getBOREevents(event):
    #valid = ["BOR", "BMOR", "HTOR", "FOR", "SEOR"]
    groups = []
    ingroup = False
    for i in range(len(df.index)):
        if (df.iat[i,8]) == event:
            #print("index: " + str(i))
            for group in groups:
                if group.groupnumber == df.iat[i,2]:
                    group.members.append(df.iat[i,3])
                    ingroup = True
            if not ingroup:
                groups.append(BOREGroup(df.iat[i,3],df.iat[i,19],df.iat[i,18],df.iat[i,2],event))
    return groups

def calcBOREfinal(groups):
    for group in groups:
        group.final = group.written + group.oral
def rankevent(groups):
    ranking = []
    for group in groups:
        if len(ranking) > 0:
            notlast = False
            for i in range(len(ranking)):
                if group.final > ranking[i].final:
                    ranking.insert(i,group)
                    notlast = True
                    break
                if not notLast:
                    ranking.append(group)
        else:
            ranking.append(group)
    return ranking
def printranking(groups):
    for i in range(len(groups)):
        print(str(i+1) + ": Group #" + str(groups[i].groupnumber) + " - Members: " + str(groups[i].members))
bore = getBOREevents("SEOR")
calcBOREfinal(bore)
printranking(rankevent(bore))
