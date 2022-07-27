import pandas as pd
groupdf = pd.read_excel("https://docs.google.com/spreadsheets/d/e/2PACX-1vS_hickoDp__qsjJ-kq1osestfyWHWKnnydQ6cZSpFj9U3Ee_iCTDTxXop0LMdaQVTE79rpYrS9U8nl/pub?output=xlsx")
def getgroupinfo():
    dict = {}
    for i in range(len(groupdf.index)):
        members = [groupdf.iat[i,1]]
        membersid = [groupdf.iat[i,2]]
        if isinstance(groupdf.iat[i,3],str):
            members.append(groupdf.iat[i,3])
            membersid.append(groupdf.iat[i,4])
        if isinstance(groupdf.iat[i,5],str):
            members.append(groupdf.iat[i,5])
            membersid.append(groupdf.iat[i,6])
        dict[groupdf.iat[i,0]] = [members,membersid]
    return dict



#puts extra info into groupdict dictionary
groupdict = getgroupinfo()

#prints entire dictionary
print(groupdict)

#prints out all information for team number 1
print(groupdict[1])

#prints out names of people in team number 2
print(groupdict[2][0])

#prints out the team captain's id for team number 2
print(groupdict[2][1][0])
