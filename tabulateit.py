from unicodedata import name
import pandas as pd
import operator

df = pd.read_excel("https://docs.google.com/spreadsheets/d/e/2PACX-1vSU3pGKAZGlgvwP0G90ZMxmbHWbhU671uxY919Nl4tzXG0-qYBewUTOZOkfkr0vsTmu9p1Szi73MlFo/pub?output=xlsx")
teamnumber = 2
studentid = 3
PBMrolePlayrow = 12 #principles of business administration score row
TTDMrolePlayrow = 13 #team decision making score row
PFLrolePlayrow =  14 #personal finance literacy score row 
IndividualSeriesrolePlayrow1 = 15 # individual series role play 1 score row
IndividualSeriesrolePlayrow2 = 16 # individuaal series role play 2 score row
BORwrittenReportrow =  17 # business operations research events written report score
BORoralPresentation = 18 # business operations research events presentation score row
######## 
EntrepreneurshipwrittenRow = 21 # entrepreneurship written report score row 
EntrepreneurshippresentationRow = 22 # entrepreneurship presentation score row
IMCEwrittenEventRow = 23 # integrated marketing written report score row 
IMCEpresentationRow = 24 # integrated marketing presentation score row 
ProfSellingpPresentScore = 25 # professional selling and consulting presentation score row

groupdf = pd.read_excel("https://docs.google.com/spreadsheets/d/e/2PACX-1vQctW2FtMUvQzUp7fiMZy5KuvijvqATJgDFlqLN97T5BhQ60jFiknOUYAdXKSJLCjCTtMpIPwJtMsmw/pub?output=xlsx")

class EventInfo():
    weightedscore = 0
    def __init__(self, name, score1, score2, examscore):
        self.name = name
        self.score1 = score1
        self.score2 = score2
        self.examscore = examscore 

class GroupMember():
    events = []  # List of EventInfo per member
    finalscore = 0
    def __init__(self, membername, memberid, teamn):
        self.membername = membername
        self.memberid = str(memberid)
        self.teamnumber = teamn
        self.events = []
        self.finalscore = 0

class Group():
    finalscore = 0
    members = []
    groupnumber = 0
    def __init__(self, teamn):
        self.groupnumber = teamn
        self.members = []
        self.finalscore = 0

# Dictionary has:
# Key: Team Number
# Value: Group Class
#    Group class has a list of Group Members (each entry is a Group Member class):
#       [GroupMember0, GroupMember1, GroupMember2]
#       GroupMember0 is the Captain of the Team
def creategroupdictionary():
    dict = {}
    for i in range(len(groupdf.index)):
        # List has [Captain_GroupMember, GroupMember0, GroupMember1]
        group = Group(groupdf.iat[i,0])
        # Create Captain Member instance of class GroupMember
        group.members.append(GroupMember(groupdf.iat[i,1], groupdf.iat[i,2], groupdf.iat[i,0]))
        if isinstance(groupdf.iat[i,3],str):
            # Create Group 1 Member instance of class GroupMember
            group.members.append(GroupMember(groupdf.iat[i,3], groupdf.iat[i,4], groupdf.iat[i,0]))
        if isinstance(groupdf.iat[i,5],str):
            # Create Group 2 Member instance of class GroupMember
            group.members.append(GroupMember(groupdf.iat[i,5], groupdf.iat[i,6], groupdf.iat[i,0]))
        dict[groupdf.iat[i,0]] = group
    return dict

examdf = pd.read_excel("https://docs.google.com/spreadsheets/d/e/2PACX-1vQNO-MV8L-40GjbXUEvP4yDdkOU7zAQNnQXC5u5F3TOArdiO5sdKUE5Z_5uehkOGSUkixaqMYenG4q5/pub?output=xlsx")
#returns exam score for a student given a student id. 
def getexam(studentid):
    for i in range(len(examdf.index)):
        studentidfromsheet = examdf.iat[i,1]
        if (str(studentidfromsheet)) == studentid:
            return (examdf.iat[i,4])
    return 0

#ALGORITMS############################################################
# Principle of Business Administration Roleplays and personal finance literacy
# role play and exam score percentage are added 
# role play is 67% of total score and exam is 33%
totalRolePlayScore = 100
 
def examAndRolePlayEvents(event, shouldrank):
    if event not in EventGroupDict:
        return []
    groups = EventGroupDict[event]
    for group in groups:
        examScoreTotal = 0
        rolePlayScore = 0
        listofeventnames = []
        for student in group.members:
            sid = student.memberid
            evlist = student.events
            for ev in evlist:
                if (ev.name != event):
                    continue
                examScore = ev.examscore * 0.33 * 100
                examScoreTotal = examScoreTotal + examScore
                if ev.name not in listofeventnames:
                    rolePlayScore = (ev.score1 / totalRolePlayScore) * 0.67 * 100
                    listofeventnames.append(ev.name)
        group.finalscore = (examScoreTotal / len(group.members))  +  rolePlayScore
    if (shouldrank):
        rankedevents = rankevent(event)
    else:
        rankedevents = []
    return rankedevents

#####################################################################
def professionalSellingAndConsulting(event, shouldrank):
    if event not in EventGroupDict:
        return []
    groups = EventGroupDict[event]
    for group in groups:
        examScoreTotal = 0
        presentation_score = 0
        listofeventnames = []
        for student in group.members:
            sid = student.memberid
            evlist = student.events
            for ev in evlist:
                if (ev.name != event):
                    continue
                examScore = ev.examscore * 0.30 * 100
                examScoreTotal = examScoreTotal + examScore
                if ev.name not in listofeventnames:
                    presentation_score = (ev.score1 / totalRolePlayScore) * 0.70 * 100
                    listofeventnames.append(ev.name)
        group.finalscore = (examScoreTotal / len(group.members))  +  presentation_score
    if (shouldrank):
        rankedevents = rankevent(event)
    else:
        rankedevents = []
    return rankedevents


def IndividualSeriesRoleplay(event, shouldrank):
    if event not in EventGroupDict:
        return []
    groups = EventGroupDict[event]
    for group in groups:
        examScoreTotal = 0
        rolePlayScore1 = 0
        rolePlayScore2 = 0
        rolePlayScore = 0
        listofeventnames = []
        for student in group.members:
            sid = student.memberid
            evlist = student.events
            for ev in evlist:
                if (ev.name != event):
                    continue
                examScore = ev.examscore * 100
                examScoreTotal = examScoreTotal + examScore
                if ev.name not in listofeventnames:
                    rolePlayScore1 = (ev.score1)
                    rolePlayScore2 = (ev.score2)
                    rolePlayScore = rolePlayScore = ((rolePlayScore1 + rolePlayScore2) / (2 * totalRolePlayScore)) * (2/3.0) * 100
                    listofeventnames.append(ev.name)
        group.finalscore = ((examScoreTotal * (1/3.0))  +  rolePlayScore)
    if (shouldrank):
        rankedevents = rankevent(event)
    else:
        rankedevents = []
    return rankedevents

def Integratedmarketing(event, shouldrank):
    if event not in EventGroupDict:
        return []
    groups = EventGroupDict[event]
    for group in groups:
        examScoreTotal = 0
        presentation_score = 0
        written_report = 0
        totalScore = 0 
        listofeventnames = []
        for student in group.members:
            sid = student.memberid
            for ev in student.events:
                if (ev.name != event):
                    continue
                examScore = ev.examscore * 100
                examScoreTotal = examScoreTotal + examScore
                if ev.name not in listofeventnames:
                    presentation_score = (ev.score1 / totalRolePlayScore) * 0.35
                    written_report = (ev.score2 / totalRolePlayScore) * 0.35
                    totalScore = (presentation_score + written_report) * 100

                    listofeventnames.append(ev.name)
        group.finalscore = ((examScoreTotal / len(group.members)) * 0.3)  +  totalScore 
    if (shouldrank):
        rankedevents = rankevent(event)
    else:
        rankedevents = []
    return rankedevents
 
def calcBOREfinal(event, shouldrank):
    if event not in EventGroupDict:
        return []
    groups = EventGroupDict[event]
    for group in groups:
        # There is only 1 member here
        listofeventnames = []
        for student in group.members:
            sid = student.memberid
            evlist = student.events
            for ev in evlist:
                if (ev.name != event):
                    continue
                if ev.name not in listofeventnames:
                    lscore1 = (ev.score1 / 60) * 0.60 * 100
                    lscore2 = (ev.score2 / 40) * 0.40 * 100
                    listofeventnames.append(ev.name)
                    group.finalscore = lscore1 + lscore2
    if (shouldrank):
        rankedevents = rankevent(event)
    else:
        rankedevents = []
    return rankedevents

def calcEEfinal(event, shouldrank):
    if event not in EventGroupDict:
        return []
    groups = EventGroupDict[event]
    for group in groups:
        # There is only 1 member here
        listofeventnames = []
        for student in group.members:
            sid = student.memberid
            evlist = student.events
            for ev in evlist:
                if (ev.name != event):
                    continue
                if ev.name not in listofeventnames:
                    if ((ev.name == "EIP") or (ev.name == "IBP")):
                        lscore1 = (ev.score1 / 90) * 0.53 * 100
                        lscore2 = (ev.score2 / 80) * 0.47 * 100
                        listofeventnames.append(ev.name)
                        group.finalscore = lscore1 + lscore2
                    elif ((ev.name == "EBG") or (ev.name == "EFB") or (ev.name == "EIB")):
                        lscore1 = (ev.score1 / 60) * 0.60 * 100
                        lscore2 = (ev.score2 / 40) * 0.40 * 100
                        listofeventnames.append(ev.name)
                        group.finalscore = lscore1 + lscore2
                    else:
                        lscore1 = (ev.score1 / 100) * 0.555 * 100
                        lscore2 = (ev.score2 / 80) * 0.445 * 100
                        listofeventnames.append(ev.name)
                        group.finalscore = lscore1 + lscore2
    if (shouldrank):
        rankedevents = rankevent(event)
    else:
        rankedevents = []
    return rankedevents

def rankevent(event):
    ranking = []
    if event not in EventGroupDict:
        return ranking
    listofgroups = EventGroupDict[event]

    # Sort list of groups for this event category
    # using the operator package imported above
    for entry in (sorted(listofgroups, key=operator.attrgetter('finalscore'))):
        # Return in descending order, for ascending order just use ranking.insert(entry)
        ranking.insert(0, entry)
    return ranking

def getAllEvents():
    for i in range(len(df.index)):
        for key in EventMappingDict:
            if (df.iat[i, key]) in EventMappingDict[key]['Events']:
                event = df.iat[i, key]
                teamid = df.iat[i, teamnumber]
                if teamid in GroupDict:
                    groupinfo = GroupDict[teamid]
                    #print("index: " + str(i))
                    # Go through each member in the group and add the even and score info
                    for m in groupinfo.members:
                        actualscore1 = 0
                        actualscore2 = 0
                        examscore = 0
                        score1 = EventMappingDict[key]['Score1Location']
                        score2 = EventMappingDict[key]['Score2Location']
                        isexam = EventMappingDict[key]['HasExam']
                        if (score1 >= 0):
                            actualscore1 = df.iat[i,score1]
                        if (score2 >= 0):
                            actualscore2 = df.iat[i,score2]
                        if (isexam != 0):
                            examscore = getexam(m.memberid)
                        instanceofevent = EventInfo(event, actualscore1, actualscore2, examscore)
                        m.events.append(instanceofevent)
                        if event not in EventGroupDict:
                            EventGroupDict[event] = [groupinfo]
                        else:
                            if (groupinfo not in EventGroupDict[event]):
                                EventGroupDict[event].append(groupinfo)

#END OF ALGORITHMS########################################################################################

###### Setup Global Dictionaries Used #######
# Group Dictionary that holds group information
GroupDict = creategroupdictionary()

# Dictionary that maps events (the event is the key) to a list of groups (the value)
EventGroupDict = {}

# Dictionary that stores output information that can be written to a sheet
OutputDataDict = {}

# Maps a google sheet event column to events, the columns that have scores and
#   if the event has an exam
# Store in EventMappingDict
EventMappingDict = {}

EventMappingDict = {
    7: {'Events': ["BOR", "BMOR", "HTOR", "FOR", "SEOR"],
        'Score1Location': BORwrittenReportrow,
        'Score2Location': BORoralPresentation,
        'HasExam': 0
    },
    9: {'Events': ["EIP","ESB","EIB","IBP","EGB","EFB"],
        'Score1Location': EntrepreneurshipwrittenRow,
        'Score2Location': EntrepreneurshippresentationRow,
        'HasExam': 0
    },
    3: {'Events': ["PBM", "PFN", "PHT", "PMK"],
        'Score1Location': PBMrolePlayrow,
        'Score2Location': -1,
        'HasExam': 1
    },
    4: {'Events': ["BLTDM", "BTDM", "ETDM", "FTDM", "HTDM", "MTDM", "STDM", "TTDM"],
        'Score1Location': TTDMrolePlayrow,
        'Score2Location': -1,
        'HasExam': 1
    },
    5: {'Events': ["PFL"],
        'Score1Location': PFLrolePlayrow,
        'Score2Location': -1,
        'HasExam': 1
    },
    11: {'Events': ["FCE", "HTPS", "PSE"],
        'Score1Location': ProfSellingpPresentScore,
        'Score2Location': -1,
        'HasExam': 1
    },
    6: {'Events': ["ACT", "AAM", "ASM", "BFS", "BSM", "ENT", "FMS", "HLM", "HRM", "MCS", "QSRM", "RFSM", "RMS", "SEM"],
        'Score1Location': IndividualSeriesrolePlayrow1,
        'Score2Location': IndividualSeriesrolePlayrow2,
        'HasExam': 1
    },
    10: {'Events': ["IMCE", "IMCP", "IMCS"],
        'Score1Location': IMCEwrittenEventRow,
        'Score2Location': IMCEpresentationRow,
        'HasExam': 1
    }
}

###### Global Dictionaries Used #######


def printplaintextgroupinfo(groups, event):
    ranknum = 1
    for group in groups:
        printstr = "\t" + str(ranknum) + " " + "Event: " + event + " "
        # Get the captain's name
        printstr = printstr + "Captain: " + group.members[0].membername + " "
        for i in range(len(group.members)):
            # Don't print captain name again
            if (i == 0):
                continue
            m = group.members[i]
            printstr = printstr + "Partner" + str(i) + ": " + m.membername + " "
        printstr = printstr + "Score: " + str(round(group.finalscore, 2))
        print(printstr)
        ranknum = ranknum + 1

# groups is a sorted list (from highest to lowest)
def printplaintextoutput(groups, event):
    print("--------" + event + "--------") 
    printplaintextgroupinfo(groups, event)

# Returns an output formatted dictionary for one event
def createoneoutputdict(groups, event):
    ranknum = 1
    finaldict = {}
    ranklist = []
    eventlist = []
    captainlist = []
    partner1list = []
    partner2list = []
    scorelist = []
    for group in groups:
        # The rank
        ranklist.append(ranknum)
        # The event
        eventlist.append(event)
        # Captain
        captainlist.append(group.members[0].membername)
        if (len(group.members) == 1):
            partner1list.append("")
            partner2list.append("")
        if (len(group.members) == 2):
            partner1list.append(group.members[1].membername)
            partner2list.append("")
        if (len(group.members) == 3):
            partner1list.append(group.members[1].membername)
            partner2list.append(group.members[2].membername)
        scorelist.append(round(group.finalscore, 2))
        ranknum = ranknum + 1
    finaldict['Ranking'] = ranklist
    finaldict['Event'] = eventlist
    finaldict['Captain Name'] = captainlist
    finaldict['Partner1 Name'] = partner1list
    finaldict['Partner2 Name'] = partner2list
    finaldict['Score'] = scorelist
    return finaldict

# Populate the output format dictionary
def createoutputdict(groups, event):
    if (len(groups) == 0):
        return
    eventdict = createoneoutputdict(groups, event)
    for key in eventdict:
        if key in OutputDataDict:
            for v in eventdict[key]:
                OutputDataDict[key].append(v)
        else:
            OutputDataDict[key] = eventdict[key]

outputfile = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTmvRNtXDDzdKQD7LPWZ3qFKv6robP0N-gtIFoO7LOCMQbDXALTwiOt-qC7hXjkbGbxgN2xug61_qjP/pub?output=xlsx"
def writeToExcel():
#    writer = pd.ExcelWriter(outputfile, engine='xlsxwriter')
    outf = pd.DataFrame(OutputDataDict)
    writer = pd.ExcelWriter('D:\\Users\\Karthik\\tabulation\\Tabulation\\foo.xlsx', engine='xlsxwriter')
    outf.to_excel(writer, sheet_name='Tabulation Test Output', index=False)
    writer.save()

# Event Processing functions
def doboreevents():
    boreevents = ["BOR", "BMOR", "HTOR", "FOR", "SEOR"]
    for event in boreevents:
        rankedevents = calcBOREfinal(event, 1)
        createoutputdict(rankedevents, event)
        printplaintextoutput(rankedevents, event)

def doeeevents():
    eeevents = ["EIP","ESB","EIB","IBP","EBG","EFB"]
    for event in eeevents:
        rankedevents = calcEEfinal(event, 1)
        createoutputdict(rankedevents, event)
        printplaintextoutput(rankedevents, event)

def dopbmevents():
    pbmeventlist = ["PBM", "PFN", "PHT", "PMK"]
    for event in pbmeventlist: 
        rankedevents = examAndRolePlayEvents(event, 1)
        createoutputdict(rankedevents, event)
        printplaintextoutput(rankedevents, event)

def dotdmevents():
    tdmeventlist = ["BLTDM", "BTDM", "ETDM", "FTDM", "HTDM", "MTDM", "STDM", "TTDM"]
    for event in tdmeventlist:
        rankedevents = examAndRolePlayEvents(event, 1)
        createoutputdict(rankedevents, event)
        printplaintextoutput(rankedevents, event)

def dopfleventlist():
    pfleventlist = ["PFL"]
    for event in pfleventlist:
        rankedevents = examAndRolePlayEvents(event, 1)
        createoutputdict(rankedevents, event)
        printplaintextoutput(rankedevents, event)

def dopsandcevents():
    psandceventlist = ["FCE", "HTPS", "PSE"]
    for event in psandceventlist:
        rankedevents = professionalSellingAndConsulting(event, 1)
        createoutputdict(rankedevents, event)
        printplaintextoutput(rankedevents, event)

def isrevents():
    isreventlist = ["ACT", "AAM", "ASM", "BFS", "BSM", "ENT", "FMS", "HLM", "HRM", "MCS", "QSRM", "RFSM", "RMS", "SEM"]
    for event in isreventlist:
        rankedevents = IndividualSeriesRoleplay(event, 1)
        createoutputdict(rankedevents, event)
        printplaintextoutput(rankedevents, event)

def imcevents():
    imceventlist = ["IMCE", "IMCP", "IMCS"]
    for event in imceventlist:
        rankedevents = Integratedmarketing(event, 1)
        createoutputdict(rankedevents, event)
        printplaintextoutput(rankedevents, event)

def processallevents():
    dopfleventlist() # personal finance literacy
    doboreevents() # business operations research
    doeeevents() # entrepreneurship events
    dopbmevents() # principles of business administration
    dotdmevents() # team decision making
    dopsandcevents() # professional selling and consulting  (correct it)
    isrevents() # individual series roleplays
    imcevents() # integrated marketing campaign events  

# Main Execution
# Get all events
getAllEvents()
# Now, process the events
processallevents()
# Write to spreadsheet
writeToExcel()
