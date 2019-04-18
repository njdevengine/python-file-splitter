import pandas as pd
path = r"S:\Directory\input.xlsx"
df = pd.read_excel(path,encoding="UTF-8")
df = df[["Headers","TO","KEEP"]]

people = df["Person"].to_list()

unique = []
for i in people:
    if i not in unique:
        unique.append(i)
        
myarray = unique_reps

dates = ["Earliest Date","Date"]

#convert lame excel datetime
for i in dates:
    df[i] = pd.to_datetime(df[i],unit="D")
    df[i] = df[i].apply(lambda x: x - pd.DateOffset(years=70,days=1))
    df[i] = df[i].dt.strftime('%m/%d/%Y')

for i in range(len(dates)):
    df.loc[df[dates[i]] == '12/31/1899', dates[i]] = ""
    
for i in range(len(myarray)):
    dff = df[df["Person"] == myarray[i]]
    dff.to_csv(str(myarray[i])+".csv",encoding = "utf-8", index =False, header = True)
    print(str(len(dff))+myarray[i])
