import pandas as pd
#identify each unique google sheets files, change this to access a new sheet
sheet_id = '1OLepef_1NoFguta9IFqEQR87uMm2hH4DXo8_iL5Ccxo'

df = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv")

#print the last 3 rows of the google sheets
print(df.tail(3))
#access columns with the same name
james = df[['score2','score3']]

print(type(james))
#access a specific row
row = df.iloc[1]
print(row)

test = df['score3']


test = df.iloc[1]
print(type(test))
#have a loop loop through all participants and create final score for them
#
#turn the row into a list for easier access of scores
test = df.iloc[1].tolist()
one  = test[1]
print(test[2]+ test[1])
print(type(test))
print(type(one))
#this will be where final score will be inserted for ranking
sample_col = [0,9,8,7]
insert(loc = 0, column = 'player')
