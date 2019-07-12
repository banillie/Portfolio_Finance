'''
Programme to compare a specific year or total wlc information. It produces a wb output with data and calculations
only i.e. no graph. The output shows changes to wlc in relation 1) overall figures, 2) change between quarters,
3) percent change are highlighted in red if change is greater/less than £100m/-£100m or percentage change greater/less
than 5%/-5% of project value
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

def get_yearly_costs(data, cost_list, year):
    output_dict = {}
    for name in data:
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

'''INSTRUCTIONS FOR RUNNING PROGRAMME'''

'''1) specify file paths to where master data for analysis is stored.'''
latest_q_data = project_data_from_master("C:\\Users\\Standalone\\Will\\masters folder\\core data\\master_4_2018.xlsx")
other_q_data = project_data_from_master("C:\\Users\\Standalone\\Will\\masters folder\\core data\\master_3_2018.xlsx")

'''2) decide which output you require'''

'''option one - in year financial information'''
'''in year cost lists is chosen through the cost list. No not change.'''
cost_list = [' RDEL Forecast Total', ' CDEL Forecast Total', ' Forecast Non-Gov']
'''in year income list is chosen through the income list. No not change.'''
income_list = [' Forecast - Income both Revenue and Capital']
'''chose financial year of interest. change accordingly'''
year_interest = '19-20'

'''get fy information by entering the appropriate variables'''
one_fy = get_yearly_costs(latest_q_data, cost_list, year_interest)
two_fy = get_yearly_costs(other_q_data, cost_list, year_interest)

'''option two - for wlc costs'''
'''chose wlc cost key of interest'''
wlc_key = 'Total Forecast - Income both Revenue and Capital'
one_wlc = get_wlc(latest_q_data, wlc_key)
two_wlc = get_wlc(other_q_data, wlc_key)

'''enter desired variables into the compare function and specify file path for where output document to be saved'''
output = compare(one_wlc, two_wlc)
output.save("C:\\Users\\Standalone\\Will\\total_income_Q4_1819_data.xlsx")