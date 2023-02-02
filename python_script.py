import pandas as pd
import pandas_profiling as pp
import statistics as s
import pandas_profiling as pp
filedat=pd.read_excel("Combined.xlsx")
df=pd.DataFrame(filedat)
profrep=pp.ProfileReport(filedat)
profrep.to_file("index.html")

openlist=[]
closelist=[]
highlist=[]
lowlist=[]
adjcloselist=[]
datelist=[]
volumelist=[]
finallist=[]
namelist= sorted(set([x.strip('.csv') for x in df['Source.Name']]))
count=0
for j in range(len(df)):
    symbol=namelist[count]
    mystr=df['Source.Name'][j]
    mystr=mystr.strip('.csv')
    if mystr!=symbol: 
        x=[symbol,str(datelist[0]).strip('00:00:00'),str(datelist[-1]).strip('00:00:00'),s.mean(highlist),s.mean(lowlist),s.mean(openlist),s.mean(closelist),s.mean(adjcloselist),s.mean(volumelist),adjcloselist[-1]-adjcloselist[0]]
        finallist.append(x)
        

        openlist=[]
        closelist=[]
        highlist=[]
        lowlist=[]
        adjcloselist=[]
        datelist=[]
        volumelist=[]
        
        count+=1
    else:
        
        highlist.append(df['High'][j])
        lowlist.append(df['Low'][j])
        openlist.append(df['Open'][j])
        closelist.append(df['Close'][j])
        datelist.append(df['Date'][j])
        adjcloselist.append(df['Adj Close'][j])
        volumelist.append(df['Volume'][j])
print(finallist)        
finalframe=pd.DataFrame(finallist,columns=['Stock','From','To','Mean(High)','Mean(Low)','Mean(Open)','Mean(Close)','Mean(Adj. Close)','Mean(Vol)','Adj.Close Change'])    
print(finalframe)
finalframe.to_excel('ChangeSheet.xlsx')    

   





           



