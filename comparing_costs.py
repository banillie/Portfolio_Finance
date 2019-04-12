'''
Programme to compare projects wlc values.

It can be used to compare a specific year or total wlc information. It produces a wb output with data and calculations
only i.e. no graph.

Changes to wlc in relation 1) overall figures, 2) change between quarters, 3) percent change are highlighted
in red if change is greater/less than £100m/-£100m or percentage change greater/less than 5%/-5% of project value
'''


from bcompiler.utils import project_data_from_master
from openpyxl import Workbook
from openpyxl.styles import Font


def compare(data_1, data_2):
    wb = Workbook()
    ws = wb.active

    for i, name in enumerate(data_1):
        '''place project names into ws'''
        ws.cell(row=i+2, column=1).value = name

        '''loop for placing wlc data into ws. highlight changes between quarters in red'''
        latest_wlc = data_1[name]
        try:
            last_wlc = data_2[name]
        except KeyError:
            last_wlc = 'None'

        ws.cell(row=i + 2, column=2).value = latest_wlc

        if latest_wlc != last_wlc:
            ws.cell(row=i + 2, column=2).font = red_text

        if name in data_2.keys():
            try:
                ws.cell(row=i + 2, column=3).value = last_wlc
                change = latest_wlc - last_wlc
                if last_wlc > 0:
                    percent_change = (latest_wlc - last_wlc)/last_wlc
                else:
                    percent_change = (latest_wlc - last_wlc)/(last_wlc + 1)
                ws.cell(row=i + 2, column=4).value = change
                ws.cell(row=i + 2, column=5).value = percent_change
                if change >= 100 or change <= -100:
                    ws.cell(row=i + 2, column=4).font = red_text
                if percent_change >= 0.05 or percent_change <= -0.05:
                    ws.cell(row=i + 2, column=5).font = red_text
            except TypeError:
                pass
        else:
            ws.cell(row=i + 2, column=3).value = last_wlc

    ws.cell(row=1, column=1).value = 'Project Name'
    ws.cell(row=1, column=2).value = 'Latest Quarter'
    ws.cell(row=1, column=3).value = 'Last Quarter'
    ws.cell(row=1, column=4).value = 'Change'
    ws.cell(row=1, column=5).value = 'Percentage Change'
    return wb

def get_yearly_costs(data, cost_list, year, remove):
    output_dict = {}
    for name in data:

        if name not in remove:
            project_dict = data[name]
            total = 0
            for type in cost_list:
                if year + type in project_dict.keys():
                    cost = project_dict[year + type]
                    try:
                        total = total + cost
                    except TypeError:
                        pass

        output_dict[name] = total

    return output_dict

def get_wlc(data, key):
    output_dict = {}
    for name in(data):
        total = data[name][key]
        output_dict[name] = total

    return output_dict

red_text = Font(color="FF0000")

'''Data'''
latest_q = project_data_from_master("C:\\Users\\Standalone\\Will\\masters folder\\master_3_2018.xlsx")
last_q = project_data_from_master("C:\\Users\\Standalone\\Will\\masters folder\\master_2_2018.xlsx")

'''option to remove projects, such as rail franchising projects'''
remove_projects = ['West Coast Partnership Franchise', 'South Eastern Rail Franchise Competition',
                   'Rail Franchising Programme', 'East Midlands Franchise', 'HS2 Phase 2b',
                   'HS2 Phase1', 'HS2 Phase2a']

'''for yearly information'''
#cost_list = [' RDEL Forecast Total', ' CDEL Forecast Total', ' Forecast Non-Gov']
#year_interest = '23-24'
#one = get_yearly_costs(latest_q, cost_list, year_interest, remove_projects)
#two = get_yearly_costs(last_q, cost_list, year_interest, remove_projects)

'''for wlc'''
wlc_key = 'Pre 18-19 Forecast Non-Gov'
one = get_wlc(latest_q, wlc_key)
two = get_wlc(last_q, wlc_key)

output = compare(one, two)

output.save("C:\\Users\\Standalone\\Will\\for_crossrail.xlsx")