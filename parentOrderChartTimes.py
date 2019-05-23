import pandas as pd
import matplotlib.pyplot as plt
import datetime as dt
import numpy as np

## Definitions
path_to_file = 'Labor Reporting_05_13_2019_082756.xlsx'

startTime = dt.datetime(2019, 5, 1, 22)
endTime = dt.datetime(2019, 5, 2, 8)

site = 'ATL'

# Shifts
first_shift_start = dt.time(7)
first_shift_end = dt.time(15)
second_shift_start = dt.time(15)
second_shift_end = dt.time(23)
third_shift_start = dt.time(23)
third_shift_end = dt.time(7)

shift_changes = [first_shift_start, second_shift_start, third_shift_start]
this_shift_changes = [dt.datetime.combine(startTime.date(), shift_change) for shift_change in shift_changes]
#print(this_shift_changes)

## Helper functions
def shortName(name):
    """ Takes a "First Last" name string and returns "First L" """
    try:
        first, last = name.split(' ')
    except ValueError:
        first, _, last = name.split(' ')

    short = "{} {}".format(first, last[0])

    return short

## read and filter labor report into pandas dataframe
df = pd.read_excel(path_to_file)
sitemask = (df['Site'] == site) # create mask for filtering to specified site
df = df.loc[sitemask] # filter to site


# deal with overlapping transactions at edges of specified time range
# case where start outside and end inside specified range
start_overlap_mask = (df['ClockIn'] < startTime) & (df['ClockOut'] > startTime)
#df2 = df.where(start_overlap_mask, (df['ClockIn'] = startTime))
start_overlap_df = df.loc[start_overlap_mask]
#print(start_overlap_df)


timemask = (df['ClockIn'] > startTime) & (df['ClockOut'] <= endTime)

# filter to target timeframe
order = df.loc[timemask] # filter to time range

# get kit number
kits = order['FGKits'].iloc[0]

# get start and end times for order
#orderstart = order['ClockIn'].min()
orderstart = order.loc[(order['Operation'] == 'CUT' + site)]['ClockIn'].min()
orderend = order.loc[(order['Operation'] == 'KIT' + site)]['ClockOut'].max()

orderlen = pd.Timedelta(endTime - startTime).total_seconds()/3600

# get list of techs
techs = order['EmpName'].unique()
num_techs = len(techs)

tech_times = []
tech_time_labels = []
names = []
lxranges = []
lcolors = []
for tech in techs:
    techtable = order.loc[order['EmpName'] == tech][['Operation','ClockIn','ClockInDurationMin']]
    names.append(shortName(tech))
    tech_time = round(techtable['ClockInDurationMin'].sum()/60,3) # total time for tech
    tech_times.append(tech_time)
    tech_pct = round(tech_time/orderlen*100,2)
    tech_time_labels.append("{} ({}%)".format(tech_time, tech_pct))
    
    lxrange = []
    color = []
    for i, row in techtable.iterrows():
        start = pd.Timedelta(row['ClockIn'] - startTime).total_seconds()/3600
        dur = row['ClockInDurationMin']/60
        op = row['Operation']

        if 'CUT' in op:
            color.append('red')
        elif 'KIT' in op:
            color.append('orange')
        else:
            color.append('blue')

        lxrange.append((start, dur))
        
    lxranges.append(lxrange)
    lcolors.append(color)


# get list of machines
machs = order['Machine'].unique()
num_machs = len(machs)

mach_times = []
mxranges = []
mcolors = []
for mach in machs:
    machtable = order.loc[order['Machine'] == mach][['Operation', 'ClockIn', 'ClockInDurationMin']]
    mach_times.append(round(machtable['ClockInDurationMin'].unique().sum()/60,3)) # total time for machine
    
    mxrange = []
    color = []
    for i, row in machtable.iterrows():
        start = pd.Timedelta(row['ClockIn'] - startTime).total_seconds()/3600
        dur = row['ClockInDurationMin']/60
        op = row['Operation']

        if 'CUT' in op:
            color.append('red')
        elif 'KIT' in op:
            color.append('orange')
        else:
            color.append('blue')

        mxrange.append((start, dur))
        
    mxranges.append(mxrange)
    mcolors.append(color)

## Define variables to draw graphs
lymax = len(techs)*7
mymax = len(machs)*7
xmax = orderlen * 1.05
xmin = orderlen * -.05

# labor
lyranges = []
for i in range(0, len(techs)):
    scale = i * 7
    lyranges.append((1+scale, 5))

lyticks = []
for lyrange in lyranges:
    lyticks.append(lyrange[0]+2.5)

# machine
myranges = []
for i in range(0, len(machs)):
    scale = i * 7
    myranges.append((1+scale, 5))

myticks = []
for myrange in myranges:
    myticks.append(myrange[0]+2.5)

### Graphing

## Create graphs 
# Two plots in one figure, one for labor, another for machine
fig, (lx, mx) = plt.subplots(2, 1, sharex = True, gridspec_kw={'height_ratios':[num_techs,num_machs]})

## Adjust graph properties
start_label = startTime.strftime('%m/%d/%y %H:%M')
end_label = endTime.strftime('%m/%d/%y %H:%M')
fig.suptitle('Between {} and {}'.format(start_label, end_label), fontsize=10)
# no spacing between subplots
fig.subplots_adjust(hspace=0)
fig.subplots_adjust(left=0.15)
fig.subplots_adjust(right=0.88)

# add alternate axes
lxa = lx.twinx()
mxa = mx.twinx()
lxb = lx.twiny()

# labor subplot
# left y axis
lx.grid(True)
lx.set_xlim(xmin, xmax)
lx.set_ylim(0, lymax)
lx.set_ylabel('Tech Name')
lx.set_yticks(lyticks)
lx.set_yticklabels(names, fontsize=6)
# right y axis
lxa.set_ylabel('{} Labor Hours'.format(round(sum(tech_times),3)))
lxa.set_ylim(0, lymax)
lxa.set_yticks(lyticks)
lxa.set_yticklabels(tech_time_labels, fontsize=6)
# top x axis
lxb.set_xlim(xmin, xmax)
lxb.set_xticks([0, orderlen])
order_start_label = startTime.strftime('%m/%d/%y\n%H:%M')
order_end_label = endTime.strftime('%m/%d/%y\n%H:%M')
lxb.set_xticklabels([order_start_label, order_end_label], fontsize=8)

# reference lines for start and end
lx.axvline(x=0, color = 'g')
mx.axvline(x=0, color = 'g')
lx.axvline(x=orderlen, color='r')
mx.axvline(x=orderlen, color = 'r')

# reference lines for shifts

#TO BE DONE

# machine subplot
# left y axis
mx.grid(True)
mx.set_ylim(0, mymax)
mx.set_ylabel('Mach Name')
mx.set_xlabel('{} Hours'.format(round(orderlen,3)))
mx.set_yticks(myticks)
mx.set_yticklabels(machs, fontsize=6)
# right y axis
mxa.set_ylabel('{} Mach Hours'.format(round(sum(mach_times),3)))
mxa.set_ylim(0,mymax)
mxa.set_yticks(myticks)
mxa.set_yticklabels(mach_times, fontsize=6)
# bottom x axis
xticks = np.arange(0, orderlen, step=2)
mx.set_xticks(xticks, minor=False)


## insert data
for i in range(0,len(techs)):
    lx.broken_barh(lxranges[i], lyranges[i], facecolors=lcolors[i])

for i in range(0, len(machs)):
    mx.broken_barh(mxranges[i], myranges[i], facecolors=mcolors[i])

plt.show()
#order_date = orderstart.strftime('%y%m%d')
#plt.savefig(order_date + '_' + parent_num, dpi=300, facecolor='w', edgecolor='w', orientation='landscape')
