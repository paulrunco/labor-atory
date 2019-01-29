import pandas as pd
import matplotlib.pyplot as plt


## Definitions
path_to_file = 'Labor Reporting_01_25_2019_083736.xlsx'
parent_num = 'P0025311'

startTime = '01-07-2019'
endTime = '01-8-2019'

site = 'DTN'

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
#order = df.loc[(df['ParentOrderNo'] == parent_num)]

# filter to target timeframe
order = df.loc[(df['ClockIn'] >= startTime) & (df['ClockOut'] <= endTime) & (df['Site'] == site)]

# get kit number
kits = order['FGKits'].iloc[0]

# get start and end times for order
orderstart = order['ClockIn'].min()
orderend = order['ClockOut'].max()

orderlen = pd.Timedelta(orderend - orderstart).total_seconds()/3600

# get list of techs
techs = order['EmpName'].unique()
num_techs = len(techs)

tech_times = []
names = []
lxranges = []
lcolors = []
for tech in techs:
    techtable = order.loc[order['EmpName'] == tech][['Operation','ClockIn','ClockInDurationMin']]
    names.append(shortName(tech))
    tech_times.append(round(techtable['ClockInDurationMin'].sum()/60,3)) # total time for tech

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

### Graphing

## Create graphs 
# Two plots in one figure, one for labor, another for machine
fig, (lx, mx) = plt.subplots(2, 1, sharex = True, gridspec_kw={'height_ratios':[num_techs,num_machs]})

## Adjust graph properties
fig.suptitle('Order {} \n {}'.format(parent_num,kits), fontsize=10)
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
lxa.set_yticklabels(tech_times, fontsize=6)
# top x axis
lxb.set_xlim(xmin, xmax)
lxb.set_xticks([0, orderlen])
lxb.set_xticklabels([orderstart.strftime('%m/%d/%y\n%H:%M'), orderend.strftime('%m/%d/%y\n%H:%M')], fontsize=8)

# reference lines for start and end
lx.axvline(x=0, color = 'g')
mx.axvline(x=0, color = 'g')
lx.axvline(x=orderlen, color='r')
mx.axvline(x=orderlen, color = 'r')

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

## insert data
for i in range(0,len(techs)):
    lx.broken_barh(lxranges[i], lyranges[i], facecolors=lcolors[i])

for i in range(0, len(machs)):
    mx.broken_barh(mxranges[i], myranges[i], facecolors=mcolors[i])

plt.show()
