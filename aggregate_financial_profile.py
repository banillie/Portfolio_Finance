'''
Programme to create a financial profile for a group of projects i.e. can produce the portfolio profile or a chosen
set of projects profile.

Outputs a workbook.

work is required to make it more generic/flexible
'''

from bcompiler.utils import project_data_from_master
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font


def financial_info(list_names, master_data, cells_to_capture):

    '''
    function that creates dictionary of tuples containing financial key:value info for each project.
    list_names = list of names - the group of projects of interest
    master_data = master data set
    cells_to_capture = lists of keys of interest
    '''

    output_dicitonary = {}

    for name in list_names:
        if name in master_data.keys():
            output_list = []
            for item in cells_to_capture:
                if item in master_data[name]:
                    if master_data[name][item] is None:
                        key = item
                        value = 0
                        tuple = (key, value)
                        output_list.append(tuple)
                    else:
                        key = item
                        value = master_data[name][item]
                        tuple = (key, value)
                        output_list.append(tuple)

            output_dicitonary[name] = output_list
        else:
            pass

    return output_dicitonary


def year_totals(list_names, cells_to_capture, fin_data):

    '''
    function  calculates the total spend each year. The yearly totals are used to calculate
    the profile graph. Returns a list.
    list_names = list of names - the group of projects of interest
    cells_to_capture = lists of keys of interest
    fin_data = dictionary of project costs created via financial info
    '''

    totals_list = []
    for i in range(0, len(cells_to_capture)):
        key = cells_to_capture[i]
        thesum = 0
        for name in list_names:
            if name in remove_from_totals:
                pass
            else:
                try:
                    thesum = thesum + fin_data[name][key]
                except TypeError:
                    pass
        totals_list.append(thesum)

    #TODO fix the message below
    #print('these project\'s costs have been removed from totals' + remove_from_totals)

    return totals_list


def likeforlike(data_1, data_2):
    '''
    small programme used to filter out projects that are not in both data sets
    :param data_1: most recent quarters data
    :param data_2: less recent quarters data
    :return: a list of projects that are in both data sets
    '''

    one = list(set(data_1) - set(data_2))
    two = list(set(data_2) - set(data_1))

    output_list = one + two

    return output_list

def place_in_excel(fin_data, totals, cells_to_capture):
    '''
    function places all data into a new workbook.
    :param fin_data: dictionary of project costs created via financial info
    :param totals: list of total costs for each year
    :param cells_to_capture: lists of keys of interest
    :return: excle workbook
    '''

    wb = Workbook()
    ws = wb.active

    for i, name in enumerate(fin_data.keys()):
        '''lists project names in row one'''
        ws.cell(row=1, column=i + 2).value = name

        '''iterates through financial dictionary - placing financial data in ws'''
        for x in range(0, len(fin_data[name])):
            ws.cell(row=x+2, column=i+2).value = fin_data[name][x][1]

    '''places totals in final column. to note because this is a list and not a dictionary as for fin_data there is 
    possibility that data could become unaligned. Whether changing the list of cells_to_capture causes them to become
    unaligned needs to be tested'''
    for i, values in enumerate(totals):
        ws.cell(row=i + 2, column=len(fin_data.keys())+2).value = values

    '''places keys into the chart in the first column'''
    for i, key in enumerate(cells_to_capture):
        ws.cell(row=i+2, column=1).value = key

    '''data for overall chart. As above because this data is in a list - possibility of it being unaligned needs 
    testing. not the best way of managing data flow, but working for now'''
    start_row = len(totals) + 8
    for x in range(0, int(len(totals) / 4)):
        ws.cell(row=start_row, column=2, value=totals[x])
        start_row += 1

    start_row = len(totals) + 8
    for x in range(int(len(totals) / 4), (int(len(totals) / 4) * 2)):
        ws.cell(row=start_row, column=3, value=totals[x])
        start_row += 1

    start_row = len(totals) + 8
    for x in range((int(len(totals) / 4) * 2), (int(len(totals) / 4) * 3)):
        ws.cell(row=start_row, column=4, value=totals[x])
        start_row += 1

    start_row = len(totals) + 8
    for x in range((int(len(totals) / 4) * 3), int(len(totals))):
        ws.cell(row=start_row, column=5, value=totals[x])
        start_row += 1

    '''code was essentially a hack'''

    start_row = len(totals) + 8
    list_of_numbers = [0, len(capture_rdel), len(capture_rdel)*2]
    total_sum = 0
    for i in range(0, len(capture_rdel)):
        for x in list_of_numbers:
            total_sum = total_sum + totals[x + i]
            ws.cell(row=start_row, column=6, value=total_sum)
        start_row += 1
        total_sum = 0

    a = len(totals) + 7
    ws.cell(row=a, column=2, value='RDEL')
    ws.cell(row=a, column=3, value='CDEL')
    ws.cell(row=a, column=4, value='Non-Gov')
    ws.cell(row=a, column=5, value='Income')
    ws.cell(row=a, column=6, value='Total')

    # ws.cell(row=a+1, column=1, value='17/18')
    ws.cell(row=a + 1, column=1, value='18/19')
    ws.cell(row=a + 2, column=1, value='19/20')
    ws.cell(row=a + 3, column=1, value='20/21')
    ws.cell(row=a + 4, column=1, value='21/22')
    ws.cell(row=a + 5, column=1, value='22/23')
    ws.cell(row=a + 6, column=1, value='23/24')
    ws.cell(row=a + 7, column=1, value='24/25')
    ws.cell(row=a + 8, column=1, value='25/26')
    ws.cell(row=a + 9, column=1, value='26/27')
    ws.cell(row=a + 10, column=1, value='27/28')
    ws.cell(row=a + 11, column=1, value='Unprofiled')

    '''this builds a very basic chart'''
    # TODO fix chart
    chart = LineChart()
    chart.title = 'Portfolio cost profile'
    chart.style = 4
    chart.x_axis.title = 'Financial Year'
    chart.y_axis.title = 'Cost (£m)'
    chart.height = 15  # default is 7.5
    chart.width = 26  # default is 15

    '''styling chart'''
    # axis titles
    font = Font(typeface='Calibri')
    size = 1200  # 12 point size
    cp = CharacterProperties(latin=font, sz=size, b=True)  # Bold
    pp = ParagraphProperties(defRPr=cp)
    rtp = RichText(p=[Paragraph(pPr=pp, endParaRPr=cp)])
    chart.x_axis.title.tx.rich.p[0].pPr = pp
    chart.y_axis.title.tx.rich.p[0].pPr = pp

    # title
    size_2 = 1400
    cp_2 = CharacterProperties(latin=font, sz=size_2, b=True)
    pp_2 = ParagraphProperties(defRPr=cp_2)
    rtp_2 = RichText(p=[Paragraph(pPr=pp_2, endParaRPr=cp_2)])
    chart.title.tx.rich.p[0].pPr = pp_2

    data = Reference(ws, min_col=2, min_row=51, max_col=5, max_row=61)
    cats = Reference(ws, min_col=1, min_row=52, max_row=61)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    s3 = chart.series[0]
    s3.graphicalProperties.line.solidFill = "36708a"  # dark blue
    s8 = chart.series[1]
    s8.graphicalProperties.line.solidFill = "68db8b"  # green
    s9 = chart.series[2]
    s9.graphicalProperties.line.solidFill = "794747"  # dark red
    s9 = chart.series[3]
    s9.graphicalProperties.line.solidFill = "73527f"  # purple

    ws.add_chart(chart, "I52")

    return wb

'''

INPUT FOR RUNNING PROGRAMME

'''

'''List of financial data keys to capture. This should be amended to years of interest'''

capture_rdel = ['18-19 RDEL Forecast Total', '19-20 RDEL Forecast Total',
                 '20-21 RDEL Forecast Total','21-22 RDEL Forecast Total','22-23 RDEL Forecast Total',
                 '23-24 RDEL Forecast Total','24-25 RDEL Forecast Total','25-26 RDEL Forecast Total',
                 '26-27 RDEL Forecast Total','27-28 RDEL Forecast Total','Unprofiled RDEL Forecast Total']


capture_cdel = ['18-19 CDEL Forecast Total','19-20 CDEL Forecast Total',
                '20-21 CDEL Forecast Total','21-22 CDEL Forecast Total',
                 '22-23 CDEL Forecast Total','23-24 CDEL Forecast Total','24-25 CDEL Forecast Total',
                 '25-26 CDEL Forecast Total','26-27 CDEL Forecast Total','27-28 CDEL Forecast Total',
                 'Unprofiled CDEL Forecast Total']

capture_ng = ['18-19 Forecast Non-Gov','19-20 Forecast Non-Gov','20-21 Forecast Non-Gov','21-22 Forecast Non-Gov',
                 '22-23 Forecast Non-Gov','23-24 Forecast Non-Gov','24-25 Forecast Non-Gov',
                 '25-26 Forecast Non-Gov','26-27 Forecast Non-Gov',
                 '27-28 Forecast Non-Gov','Unprofiled Forecast-Gov']

capture_income =['18-19 Forecast - Income both Revenue and Capital', '19-20 Forecast - Income both Revenue and Capital',
                '20-21 Forecast - Income both Revenue and Capital', '21-22 Forecast - Income both Revenue and Capital',
                '22-23 Forecast - Income both Revenue and Capital', '23-24 Forecast - Income both Revenue and Capital',
                '24-25 Forecast - Income both Revenue and Capital', '25-26 Forecast - Income both Revenue and Capital',
                '26-27 Forecast - Income both Revenue and Capital', '27-28 Forecast - Income both Revenue and Capital',
                'Unprofiled Forecast Income']

all_data_lists = capture_rdel + capture_cdel + capture_ng + capture_income

''' ONE: master data to be used for analysis'''

latest_q = project_data_from_master("C:\\Users\\Standalone\\Will\\masters folder\\core data\\merged_master"
                                    "_testing.xlsx")
other_q_data = project_data_from_master("C:\\Users\\Standalone\\Will\\masters folder\\core data\\master_2_2018.xlsx")

''' TWO: project name list options - this is where the group of interest is specified '''

'''option 1 - all '''
proj_names_all = list(latest_q.keys())

'''option 2 - a group'''
#TODO write function for filtering list of project names based on group
#proj_names_group

'''option 3 - bespoke list of projects'''
#proj_names_bespoke = ['Digital Railway']

'''THREE: project names to be removed from total figures - chose as necessary'''
dont_double_count = ['HS2 Phase 2b', 'HS2 Phase1', 'HS2 Phase2a', 'East Midlands Franchise',
                      'South Eastern Rail Franchise Competition', 'West Coast Partnership Franchise']

like_for_like = likeforlike(latest_q, other_q_data)

remove_from_totals = dont_double_count + like_for_like

'''Four: enter variables created via options above into functions and run programme'''

finance_data = financial_info(proj_names_all, latest_q, all_data_lists)
total_data = year_totals(proj_names_all, all_data_lists, latest_q)

output = place_in_excel(finance_data, total_data, all_data_lists)

output.save("C:\\Users\\Standalone\\Will\\test_output.xlsx")