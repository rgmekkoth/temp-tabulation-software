

import pandas as pd
df = pd.read_excel("https://docs.google.com/spreadsheets/d/e/2PACX-1vSAH0q4IKYBmA46IT2IbrjpuxwwmOlKUcAZSaaFWa_iiv0Iv6sER08JDWAy0wJ8mjdmobE9OORIyjft/pub?output=xlsx")
class BOREGroup():
    def __init__(self, member1, oral, written, groupnumber, eventtype):
        self.members = [member1]
        self.groupnumber = groupnumber
        self.written = written
        self.oral = oral
        self.eventtype = eventtype
class EEGroup():
    def __init__(self, member1, oral, written, groupnumber, eventtype):
        self.members = [member1]
        self.groupnumber = groupnumber
        self.written = written
        self.oral = oral
        self.eventtype = eventtype
#Returns a list of all competitors for a given event for any event in the BORE category
def getBOREevents(event):
    #valid = ["BOR", "BMOR", "HTOR", "FOR", "SEOR"]
    groups = []

    for i in range(len(df.index)):
        if (df.iat[i,8]) == event:
            ingroup = False
            #print("index: " + str(i))
            for group in groups:
                if group.groupnumber == df.iat[i,2]:
                    group.members.append(df.iat[i,3])
                    ingroup = True
            if not ingroup:
                groups.append(BOREGroup(df.iat[i,3],df.iat[i,19],df.iat[i,18],df.iat[i,2],event))
    return groups
#Returns a list of all competitors for a given event for any event in the EE category
def getEEevents(event):
    groups = []
    notlast = False
    for i in range(len(df.index)):
        if (df.iat[i,10]) == event:
            ingroup = False
            for group in groups:
                if group.groupnumber == df.iat[i,2]:
                    group.members.append(df.iat[i,3])
                    ingroup = True
            if not ingroup:
                groups.append(EEGroup(df.iat[i,3],df.iat[i,23],df.iat[i,22],df.iat[i,2],event))
    return groups
#Calculates the final score for a bore event
def calcBOREfinal(groups):
    for group in groups:
        group.final = group.written + group.oral
#Calculates the final score for an EE event
def calcEEfinal(groups):
    for group in groups:
        group.final = 100 * (((group.written/100) + (group.oral/100))/2)
#Takes in a list of events, then sorts it by the final attribute, and outputs the list.
def rankevent(groups):
    ranking = []
    notlast = False
    for group in groups:
        if len(ranking) > 0:
            notlast = False
            for i in range(len(ranking)):
                if group.final > ranking[i].final:
                    ranking.insert(i,group)
                    notlast = True
                    break
                if not notlast:
                    ranking.append(group)
        else:
            ranking.append(group)
    return ranking
#prints out the ranking for a given event, inputs sorted list
def printranking(groups,event):
    print("--------" + event + "--------")
    for i in range(len(groups)):
        print(str(i+1) + ": Group #" + str(groups[i].groupnumber) + " - Members: " + str(groups[i].members) + " - Score: " + str(round(groups[i].final, 2)))


boreevents = ["BOR", "BMOR", "HTOR", "FOR", "SEOR"]
eeevents = ["EIP","ESB","EIB","IBP","EGB","EFB"]
for event in boreevents:
    eventlist = getBOREevents(event)
    calcBOREfinal(eventlist)
    printranking(rankevent(eventlist),event)
for event in eeevents:
    eventlist = getEEevents(event)
    calcEEfinal(eventlist)
    printranking(rankevent(eventlist),event)
