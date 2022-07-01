
import pandas as pd
df = pd.read_excel("https://docs.google.com/spreadsheets/d/e/2PACX-1vSAH0q4IKYBmA46IT2IbrjpuxwwmOlKUcAZSaaFWa_iiv0Iv6sER08JDWAy0wJ8mjdmobE9OORIyjft/pub?output=xlsx")
teamnumber = 2
studentid = 3
class Group():
    def __init__(self, member1, score1, score2, groupnumber, eventtype):
        self.members = [member1]
        self.groupnumber = groupnumber
        self.score1 = score1
        self.score2 = score2
        self.eventtype = eventtype

#Returns a list of all competitors for a given event for any event
#PARAMETERS
#event - string representing the abbreviation matching the event
#eventindex - integer matching the column in the spreadsheet that contains the event abbreviation
#score1 - integer matching the column in the spreadsheet where the first score is located
#score2 - integer matching the column in the spreadhseet where the second score is located
#You can access  the scores of a given group using group.score1 and groupe.score2
#This function requires the group class above.

def getevents(event,eventindex,score1=-1,score2=-1):
    #valid = ["BOR", "BMOR", "HTOR", "FOR", "SEOR"]
    groups = []
    for i in range(len(df.index)):
        if (df.iat[i,eventindex]) == event:
            ingroup = False
            #print("index: " + str(i))
            for group in groups:
                if group.groupnumber == df.iat[i,teamnumber]:
                    group.members.append(df.iat[i,studentid])
                    ingroup = True
            if not ingroup:
                groups.append(Group(df.iat[i,studentid],df.iat[i,score1],df.iat[i,score2],df.iat[i,teamnumber],event))
    return groups

#Calculates the final score for a bore event
def calcBOREfinal(groups):
    for group in groups:
        group.final = group.score1 + group.score2
#Calculates the final score for an EE event
def calcEEfinal(groups):
    for group in groups:
        group.final = 100 * (((group.score1/100) + (group.score2/100))/2)
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
    eventlist = getevents(event,8,18,19)
    calcBOREfinal(eventlist)
    printranking(rankevent(eventlist),event)
for event in eeevents:
    eventlist = getevents(event,10,22,23)
    calcEEfinal(eventlist)
    printranking(rankevent(eventlist),event)
