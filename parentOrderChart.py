import pandas as pd
import matplotlib.pyplot as plt
import datetime as dt

## Definitions
path_to_file = 'Labor Reporting_05_22_2019_085819.xlsx'
parent_num = 'P0042060'

site = 'MTP'

## Helper functions
def shortName(name):
    """ Takes a "First Last" name string and returns "First L" """
    try:
        first, last = name.split(' ')
    except ValueError:
        first, _, last = name.split(' ')

    short = "{} {}".format(first, last[0])

    return short

## read labor report into pandas dataframe
df = pd.read_excel(path_to_file)

# filter to target parent order
order = df.loc[(df['ParentOrderNo'] == parent_num)]

# filter to target timeframe

#order = df.loc[(df['ClockIn'] > startTime) & (df['ClockOut'] <= endTime) & (df['Site'] == site)]
# get kit number
kits = order['FGKits'].iloc[0]

# get start and end times for order
#orderstart = order['ClockIn'].min()
orderstart = order.loc[(order['Operation'] == 'CUT' + site)]['ClockIn'].min()
orderend = order.loc[(order['Operation'] == 'KIT' + site)]['ClockOut'].max()

orderlen = pd.Timedelta(orderend - orderstart).total_seconds()/3600

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
        start = pd.Timedelta(row['ClockIn'] - orderstart).total_seconds()/3600
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
        start = pd.Timedelta(row['ClockIn'] - orderstart).total_seconds()/3600
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

# tick to show hours of the day
hour_ticks = []
hour_ticks_labels = []
for i in range(0, int(orderlen) + 1):
    hour = dt.datetime(orderstart.year, orderstart.month, orderstart.day, orderstart.hour) + dt.timedelta(hours=1+i)
    delta = pd.Timedelta(hour - orderstart).total_seconds()/3600
    hour_ticks.append(delta)
    hour_ticks_labels.append(hour.strftime('%H:%M'))

### Graphing

## Create graphs 
# Two plots in one figure, one for labor, another for machine
fig, (labor_plot, machine_plot) = plt.subplots(2, 1, sharex = True, gridspec_kw={'height_ratios':[num_techs,num_machs]})

## Adjust graph properties
fig.suptitle('Order {} \n {}'.format(parent_num,kits), fontsize=10)
# no spacing between subplots
fig.subplots_adjust(hspace=0)
fig.subplots_adjust(left=0.15)
fig.subplots_adjust(right=0.88)

# add alternate axes
lxa = labor_plot.twinx()
mxa = machine_plot.twinx()
lxb = labor_plot.twiny()

# labor subplot
# left y axis
labor_plot.grid(axis='y')
labor_plot.set_xlim(xmin, xmax)
labor_plot.set_ylim(0, lymax)
labor_plot.set_ylabel('Tech Name')
labor_plot.set_yticks(lyticks)
labor_plot.set_yticklabels(names, fontsize=6)
# right y axis
lxa.set_ylabel('{} Labor Hours'.format(round(sum(tech_times),3)))
lxa.set_ylim(0, lymax)
lxa.set_yticks(lyticks)
lxa.set_yticklabels(tech_time_labels, fontsize=6)
# top x axis
lxb.set_xlim(xmin, xmax)
lxb.set_xticks([0, orderlen])
order_start_label = orderstart.strftime('%m/%d/%y\n%H:%M')
order_end_label = orderend.strftime('%m/%d/%y\n%H:%M')
lxb.set_xticklabels([order_start_label, order_end_label], fontsize=8)
# bottom x axis
labor_plot.grid(axis='x')
machine_plot.grid(axis='x')
labor_plot.set_xticks(hour_ticks)
labor_plot.set_xticklabels(hour_ticks_labels)

# reference lines for start and end
labor_plot.axvline(x=0, color = 'g')
machine_plot.axvline(x=0, color = 'g')
labor_plot.axvline(x=orderlen, color='r')
machine_plot.axvline(x=orderlen, color = 'r')

# reference lines for shifts
#TO BE DONE

# machine subplot
# left y axis
machine_plot.grid(axis='y')
machine_plot.set_ylim(0, mymax)
machine_plot.set_ylabel('Mach Name')
machine_plot.set_xlabel('{} Hours'.format(round(orderlen,3)))
machine_plot.set_yticks(myticks)
machine_plot.set_yticklabels(machs, fontsize=6)
# right y axis
mxa.set_ylabel('{} Mach Hours'.format(round(sum(mach_times),3)))
mxa.set_ylim(0,mymax)
mxa.set_yticks(myticks)
mxa.set_yticklabels(mach_times, fontsize=6)
# bottom x axis


## insert data
for i in range(0,len(techs)):
    labor_plot.broken_barh(lxranges[i], lyranges[i], facecolors=lcolors[i])

for i in range(0, len(machs)):
    machine_plot.broken_barh(mxranges[i], myranges[i], facecolors=mcolors[i])

plt.show()
