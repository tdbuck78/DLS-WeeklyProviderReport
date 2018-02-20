''' 

	----------WEEKLY PROVIDER REPORT---------------------

Project for Developmental Learning Solutions inc.  usage
  
  The Weekly Provider Report program is a management tool to view provider/client activity over a 4 week period
  
  This program
  
  (1) Scrapes data from an html table produced from a sql query ran on the dls server   (github: table data was saved to csv, and censored for privacy)
  
  (2) Organizes and inputs data into excel workbook with useful format  (Worksheet per provider)
  
  (3) Creates a visual for a quick readable analysis of Provider hours of last 4 weeks
  
  

'''


'''importing and scrubbing data -------------------------------'''

import pandas as pd
import time
import datetime
from datetime import date
import numpy as np
import xlsxwriter
data = pd.read_csv('weeklyreportdata.csv')

data['Hours'] = data['Hours'].apply(pd.to_numeric)
data['Start'] = data['Start'].apply(pd.to_numeric)
data['End'] = data['End'].apply(pd.to_numeric)
data['Hours'] = (round(data['Hours']/3600, 2))




import time
import datetime
from datetime import date
import numpy as np
import xlsxwriter


'''provider baselines-----------------------------------------------------------------'''

# these baselines are can be updated when necessary

baselines = {'Provider1':12, 'Provider3':1, 'Provider4':10,
             'Provider5':12,'Provider6':10, 'Provider7':8,'Provider10':8,'Provider11':18,
             'Provider12':3,'Provider13':12,'Provider14':10,'Provider15':8,'Provider16':4,
             'Provider17':1,'Provider18':22,'Provider19':22,'Provider20':22 }
             
'''making array of unix dates---------------------------------------------------------'''
currWeekNum = date.today().isocalendar()[1]
currYrNum = date.today().isocalendar()[0]
d = "1/1/{}".format(str(currYrNum))

startUnix = time.mktime(datetime.datetime.strptime(d, "%d/%m/%Y").timetuple())

endUnix = startUnix
for i in range(0,currWeekNum+1):
    endUnix += 604800

weekArr = np.linspace(endUnix-(5*604800),endUnix,6)


'''getting all table information---------------------------------------------------------'''
provArr = data['Provider'].unique()


provDic = {}
for p in provArr:
    clientArr = data[data['Provider'] == p]['Client'].unique()
    clientDic = {}
    for c in clientArr:
        weekDic = {}
        for i in range(0,5):
            lst = data[
                        (data['Provider']==p) & 
                        (data['Client']==c) & 
                        (data['Start']>weekArr[i]) & 
                        (data['Start']<weekArr[i+1])
                        ]['Hours'].tolist() 
            weekDic["Week{}".format(str(i+1))] = lst
            total = sum(lst)
        clientDic[c] = weekDic
    provDic[p] = clientDic
    
'''getting provider weekly sums---------------------------------------------------------'''
provSums = {}
provConf = {}

for p in provArr:
    totalSum = {}
    confSum = {}
    
    for i in range(0,5):
        total = data[
                    (data['Provider']==p) & 
                    (data['Start']>weekArr[i]) & 
                    (data['Start']<weekArr[i+1])
                    ]['Hours'].sum()
        
        conf = data[
                    (data['Provider']==p) & 
                    (data['Start']>weekArr[i]) & 
                    (data['Start']<weekArr[i+1])&
                    (data['Status']=="Confirmed (Appt has happened)")
                    ]['Hours'].sum()
        totalSum["Week{}".format(str(i+1))] = total
        confSum["Week{}".format(str(i+1))] = conf
    provSums[p] = totalSum
    provConf[p] = confSum
    
    


'''writing to excel---------------------------------------------------------'''
t = 'Weekly Provider Report {}.xlsx'.format(str(date.today()))
wb = xlsxwriter.Workbook(t)



borderOdd = wb.add_format({
    'border':1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color' : '#b3b3b3'
})

borderEven = wb.add_format({
    'border':1,
    'align': 'center',
    'valign': 'vcenter',
})

titleFormat = wb.add_format({                              #formats
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
})
    
weekOddFormat = wb.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color' : '#b3b3b3'
})
weekEvenFormat = wb.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
})

# adding data to worksheet
for p in provArr:
    ws = wb.add_worksheet(p)
    clientArr = data[data['Provider'] == p]['Client'].unique()
    i = 0
    for c in clientArr:
        ws.write(i+2,0,c)
        for j in range(1,6):
            wk = "Week{}".format(str(j))
            s = ""
            for k in range(0,len(provDic[p][c][wk])):
                s += "({})".format(str(provDic[p][c][wk][k]))           
            ws.write(i+2,(j*2)-1,s)
            ws.merge_range(i+2,(j*2)-1,i+2,j*2,s)
        i+=1
    for j in range(1,6):
        wk = "Week{}".format(str(j))
        if j % 2 == 0:
            ws.write(len(clientArr)+4,(j*2)-1,provSums[p][wk],borderEven)
            ws.write(len(clientArr)+3,(j*2)-1,'All Appts',borderEven)
            ws.write(len(clientArr)+4,(j*2),provConf[p][wk],borderEven)
            ws.write(len(clientArr)+3,(j*2),'Confirmed',borderEven)
        else:
            ws.write(len(clientArr)+4,(j*2)-1,provSums[p][wk],borderOdd)
            ws.write(len(clientArr)+3,(j*2)-1,'All Appts',borderOdd)
            ws.write(len(clientArr)+4,(j*2),provConf[p][wk],borderOdd)
            ws.write(len(clientArr)+3,(j*2),'Confirmed',borderOdd)

        ws.write(j,16,provSums[p][wk])
        ws.write(j,17,provConf[p][wk])
        if p in baselines.keys():
            ws.write(j,18,baselines[p])
        

    ws.merge_range('B1:K1', 'Week', titleFormat)
    ws.set_row(1,20)
    ws.set_row(0,20)
    ws.set_column(0,0,18)
    ws.set_column(2,2,10)
    ws.set_column(4,4,10)
    ws.set_column(6,6,10)                                        #more formatting
    ws.set_column(8,8,10)
    ws.set_column(10,10,10)
    ws.write('A2', 'Client', weekEvenFormat)
    ws.merge_range('B2:C2', 'Three', weekOddFormat)
    ws.write("P2",'Three')
    ws.merge_range('D2:E2', 'Two', weekEvenFormat)
    ws.write("P3",'Two')
    ws.merge_range('F2:G2', 'Last', weekOddFormat)
    ws.write("P4",'Last')
    ws.merge_range('H2:I2', 'This', weekEvenFormat)
    ws.write("P5",'This')
    ws.merge_range('J2:K2', 'Next', weekOddFormat)
    ws.write("P6",'Next')

    
    
    chart = wb.add_chart({'type': 'column'})
    line_chart = wb.add_chart({'type': 'line'})
    
    chart.add_series({
        'name':       'All Appts',
        'categories' : '={}!$P$2:$P$6'.format(p),
        'values': '={}!$Q$2:$Q$6'.format(p),
        'fill':   {'color': '#4286f4'}
    })
    
    chart.add_series({
        'name':       'Confirmed',
        'values': '={}!$R$2:$R$6'.format(p),
        'fill':   {'color': '#ff3333'}
    })
    line_chart.add_series({
        'name':       'Base',
        'values':     '={}!$S$2:$S$6'.format(p),
        'line': {
            'color': '#000000',
            'width': 3,
    }})
    
    chart.set_x_axis({
    'name': 'Week',
    'name_font': {'size': 14, 'bold': True}
    })
    
    chart.set_y_axis({
    'name': 'Hours',
    'name_font': {'size': 14, 'bold': True}
    })

    
    chart.combine(line_chart)
    ws.insert_chart("C{}".format(len(clientArr)+7), chart)




wb.close()



