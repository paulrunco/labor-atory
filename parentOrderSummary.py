import pandas as pd
import matplotlib.pyplot as plt


## Definitions
path_to_file = 'Labor Reporting_01_23_2019_163003.xlsx'

## read labor report into pandas dataframe
df = pd.read_excel(path_to_file)

# list of unique orders
orders = df['ParentOrderNo'].unique()


# initialize storage arrays

site = []
num_techs = []
num_machs = []
tot_labor_hours = []
tot_mach_hours = []
part_nums = []
start_time = []
end_time = []
duration = []

for order in orders:
    orderdf = df.loc[df['ParentOrderNo'] == order]
    #site
    site.append(orderdf['Site'].iloc[0])
    # number of technicians
    num_techs.append(len(orderdf['EmpName'].unique()))
    #number of machines
    num_machs.append(len(orderdf['Machine'].unique()))
    #order start time
    start = orderdf['ClockIn'].min()
    start_time.append(start)
    #order end time
    end = orderdf['ClockOut'].max()
    end_time.append(end)
    #order duration
    orderlen = pd.Timedelta(end - start).total_seconds()/3600
    duration.append(orderlen)
    #kit part numbers
    part = orderdf['FGKits'].iloc[0]
    part_nums.append(part)
    #total labor hours
    tot_labor_hours.append(orderdf['TotalHours'].sum())
    #total machine hours
    tot_mach_hours.append(orderdf['TotalHours'].unique().sum())
    

## Create dataframe of results
dictout = {'Site':site, 'ParentOrderNo':orders, 'TotalLaborHours':tot_labor_hours, 'TotalMachineHours':tot_mach_hours, 'FGKits':part_nums, 'StartTime':start_time, 'NumTechs':num_techs, 'NumMachs':num_machs, 'EndTime':end_time, 'DurationHours':duration}
dfout = pd.DataFrame(data=dictout)

# Reorganize column order
dfout = dfout[['Site', 'ParentOrderNo', 'FGKits', 'StartTime', 'EndTime', 'DurationHours', 'NumTechs', 'TotalLaborHours', 'NumMachs', 'TotalMachineHours']]

## Write to excel
writer = pd.ExcelWriter('ParentOrderSummary.xlsx', engine='xlsxwriter')
dfout.to_excel(writer)
writer.save()



### filter to target parent order
##order = df.loc[(df['ParentOrderNo'] == parent_num)]
##
### get start and end times for order
##orderstart = order['ClockIn'].min()
##orderend = order['ClockOut'].max()
##
##orderlen = pd.Timedelta(orderend - orderstart).total_seconds()/3600
##
### get list of techs
##techs = order['EmpName'].unique()
##num_techs = len(techs)
##
##tech_times = []
##names = []
##lxranges = []
##lcolors = []
##for tech in techs:
##    techtable = order.loc[order['EmpName'] == tech][['Operation','ClockIn','ClockInDurationMin']]
##    names.append(shortName(tech))
##    tech_times.append(round(techtable['ClockInDurationMin'].sum()/60,3)) # total time for tech
##
##    lxrange = []
##    color = []
##    for i, row in techtable.iterrows():
##        start = pd.Timedelta(row['ClockIn'] - orderstart).total_seconds()/3600
##        dur = row['ClockInDurationMin']/60
##        op = row['Operation']
##
##        if 'CUT' in op:
##            color.append('red')
##        elif 'KIT' in op:
##            color.append('orange')
##        else:
##            color.append('blue')
##
##        lxrange.append((start, dur))
##        
##    lxranges.append(lxrange)
##    lcolors.append(color)
##
##
### get list of machines
##machs = order['Machine'].unique()
##num_machs = len(machs)
##
##mach_times = []
##mxranges = []
##mcolors = []
##for mach in machs:
##    machtable = order.loc[order['Machine'] == mach][['Operation', 'ClockIn', 'ClockInDurationMin']]
##    mach_times.append(round(machtable['ClockInDurationMin'].unique().sum()/60,3)) # total time for machine
##    
##    mxrange = []
##    color = []
##    for i, row in machtable.iterrows():
##        start = pd.Timedelta(row['ClockIn'] - orderstart).total_seconds()/3600
##        dur = row['ClockInDurationMin']/60
##        op = row['Operation']
##
##        if 'CUT' in op:
##            color.append('red')
##        elif 'KIT' in op:
##            color.append('orange')
##        else:
##            color.append('blue')
##
##        mxrange.append((start, dur))
##        
##    mxranges.append(mxrange)
##    mcolors.append(color)
##
##
##
