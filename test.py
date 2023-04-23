import pandas as pd #pip install pandas

data = pd.read_excel("WithEmail.xlsx")
#print(data["City"])

if "Email" in data.columns:
    emails=list(data["Email"])
    #print(emails)
    c=[]
    for i in emails:
        #print(i)
        if pd.isnull(i) == False:
            #print(i)
            c.append(i)
    emails=c
    print(emails)
else:
    print("No exist")

