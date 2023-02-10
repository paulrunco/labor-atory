import pandas as pd

def summarize_orders(path, save_as):

    ## read labor report into pandas dataframe
    df = pd.read_excel(path)

    # list of unique orders
    orders = df["ParentOrderNo"].unique()

    # initialize storage arrays
    labor_dates = []
    sites = []
    num_techs = []
    num_machs = []
    tot_labor_hours = []
    tot_mach_hours = []
    part_nums = []
    start_time = []
    end_time = []
    duration = []

    for order in orders:
        # create dataframe for each order
        orderdf = df.loc[df["ParentOrderNo"] == order]

        # labor date
        labor_date = orderdf["LaborDate"].min()
        labor_dates.append(labor_date)

        # site
        site = orderdf["Site"].iloc[0]
        sites.append(site)

        # number of technicians
        num_techs.append(len(orderdf["EmpName"].unique()))

        # number of machines
        num_machs.append(len(orderdf["Machine"].unique()))

        # order start time
        start = orderdf["ClockIn"].min()
        start_time.append(start)

        # order end time
        last_event = 'KIT' + site
        end = orderdf.loc[(orderdf["Operation"] == last_event)]["ClockOut"].max()
        end_time.append(end)

        # order duration
        orderlen = pd.Timedelta(end - start).total_seconds() / 3600
        duration.append(orderlen)

        # kit part numbers
        part = orderdf["FGKits"].iloc[0]
        part_nums.append(part)

        # total labor hours
        tot_labor_hours.append(orderdf["TotalHours"].sum())

        # total machine hours
        cut_hours = (
            orderdf.loc[(orderdf["Operation"] == "CUT" + site)]["TotalHours"]
            .unique()
            .sum()
        )
        tot_mach_hours.append(cut_hours)

    ## Create dataframe of results
    dictout = {
        "ParentOrderNo": orders,
        "Site": sites,
        "LaborDate": labor_dates,
        "FGKits": part_nums,
        "StartTime": start_time,
        "EndTime": end_time,
        "DurationHours": duration,
        "TotalLaborHours": tot_labor_hours,
        "TotalMachineHours": tot_mach_hours,
        "NumTechs": num_techs,
        "NumMachines": num_machs,
    }

    dfout = pd.DataFrame(data=dictout)

    ## Write to excel
    with pd.ExcelWriter(save_as) as writer:
        dfout.to_excel(writer, index=False)

