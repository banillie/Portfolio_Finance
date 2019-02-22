'''
Programme to create a financial profile for individual projects.

Outputs a workbook which includes a graph.

work is required to make it more generic/flexible
'''


from openpyxl import Workbook
from bcompiler.utils import project_data_from_master
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font
#from aggregate_financial_profile import financial_info

# TODO change/tweak how it is designed so can use functions from aggregate_financial_profile

def financial_dict(name_list, master_data, cells_to_capture):
    '''
    the creation of a mini dictionary containing financial information. This is done via two functions; this one
    and the one title financial_info ( directly below ).
    :param name_list: list of project names
    :param master_data: master data for quarter of interest
    :param cells_to_capture: financial info key names. see lists below
    :return:
    '''

    output_dict = {}
    for name in name_list:
        get_dict_info = financial_info(name, master_data, cells_to_capture)
        output_dict[name] = get_dict_info

    return output_dict


def financial_info(name, master_data, cells_to_capture):
    '''
    function that creates dictionary containing financial {key : value} information.
    names = name of project
    master_data = master data set
    cells_to_capture = lists of keys of interest
    '''

    output_dict = {}

    if name in master_data.keys():
        for item in master_data[name]:
            if item in cells_to_capture:
                if master_data[name][item] is None:
                    output_dict[item] = 0
                else:
                    value = master_data[name][item]
                    output_dict[item] = value

    else:
        for item in cells_to_capture:
            output_dict[item] = 0

    return output_dict

def calculate_totals(name, fin_data):
    '''
    :param name: project name
    :param fin_data: mini project financial dictionary
    :return: a total number for rdel, cdel, and ng spend each financial year
    '''

    working_data = fin_data[name]
    rdel_list = []
    cdel_list = []
    ng_list = []

    for rdel in capture_rdel:
        rdel_list.append(working_data[rdel])
    for cdel in capture_cdel:
        cdel_list.append(working_data[cdel])
    for ng in capture_ng:
        ng_list.append(working_data[ng])

    total_list = []
    for i in range(len(rdel_list)):
        total = rdel_list[i] + cdel_list[i] + ng_list[i]
        total_list.append(total)

    return total_list

def calculate_income_totals(name, fin_data):
    '''
    :param name: project name
    :param fin_data: mini project financial dictionary
    :return: a total number for income spending each financial year
    '''

    working_data = fin_data[name]
    income_list = []

    for income in capture_income:
        income_list.append(working_data[income])

    return income_list

def place_in_excel(name, latest_fin_data, last_fin_data, baseline_fin_data):
    '''
    function places all data into excel spreadsheet and creates chart.
    data is placed into sheet in reverse order (see how data_list is ordered) so that most recent
    data is displayed on right hand side of the data table
    '''

    wb = Workbook()
    ws = wb.active
    data_list = [baseline_fin_data, last_fin_data, latest_fin_data]
    count = 0

    '''places in raw/reported data'''
    for data in data_list:
        for i, key in enumerate(capture_rdel):
            ws.cell(row=i+3, column=2+count, value=data[name][key])
        for i, key in enumerate(capture_cdel):
            ws.cell(row=i+3, column=3+count, value=data[name][key])
        for i, key in enumerate(capture_ng):
            ws.cell(row=i+3, column=4+count, value=data[name][key])
        count += 4

    '''places in totals'''
    baseline_totals = calculate_totals(name, baseline_fin_data)
    last_q_totals = calculate_totals(name, last_fin_data)
    latest_q_totals = calculate_totals(name, latest_fin_data)

    total_list = [baseline_totals, last_q_totals, latest_q_totals]

    c = 0
    for l in total_list:
        for i, total in enumerate(l):
            ws.cell(row=i+3, column=5+c, value=total)
        c += 4

    '''labeling data in table'''

    labeling_list_quarter = ['Baseline', 'Last Quarter', 'Latest quarter']

    ws.cell(row=1, column=2, value=labeling_list_quarter[0])
    ws.cell(row=1, column=6, value=labeling_list_quarter[1])
    ws.cell(row=1, column=10, value=labeling_list_quarter[2])

    labeling_list_type = ['RDEL', 'CDEL', 'Non-Gov', 'Total']
    repeat = 3
    c = 0
    while repeat > 0:
        for i, label in enumerate(labeling_list_type):
            ws.cell(row=2, column=2+i+c, value=label)
        c += 4
        repeat -= 1

    labeling_list_year = ['Spend', '18/19', '19/20', '20/21', '21/22', '22/23', '23/24', '24/25', '25/26', '26/27',
                          '27/28', 'Unprofiled']

    for i, label in enumerate(labeling_list_year):
        ws.cell(row=2+i, column=1, value=label)

    '''process for showing total cost profile. starting with data'''

    row_start = 16
    for x, l in enumerate(total_list):
        for i, total in enumerate(l):
            ws.cell(row=i + row_start, column=x + 2, value=total)

    '''data for graph labeling'''

    for i, quarter in enumerate(labeling_list_quarter):
        ws.cell(row=15, column=i + 2, value=quarter)

    for i, label in enumerate(labeling_list_year):
        ws.cell(row=15+i, column=1, value=label)


    chart = LineChart()
    chart.title = str(name) + ' Cost Profile'
    chart.style = 4
    chart.x_axis.title = 'Financial Year'
    chart.y_axis.title = 'Cost £m'

    '''styling chart'''
    # axis titles
    font = Font(typeface='Calibri')
    size = 1200  # 12 point size
    cp = CharacterProperties(latin=font, sz=size, b=True)  # Bold
    pp = ParagraphProperties(defRPr=cp)
    rtp = RichText(p=[Paragraph(pPr=pp, endParaRPr=cp)])
    chart.x_axis.title.tx.rich.p[0].pPr = pp
    chart.y_axis.title.tx.rich.p[0].pPr = pp
    # chart.title.tx.rich.p[0].pPr = pp

    # title
    size_2 = 1400
    cp_2 = CharacterProperties(latin=font, sz=size_2, b=True)
    pp_2 = ParagraphProperties(defRPr=cp_2)
    rtp_2 = RichText(p=[Paragraph(pPr=pp_2, endParaRPr=cp_2)])
    chart.title.tx.rich.p[0].pPr = pp_2

    '''unprofiled costs not included in the chart'''
    data = Reference(ws, min_col=2, min_row=15, max_col=4, max_row=25)
    chart.add_data(data, titles_from_data=True)
    cats = Reference(ws, min_col=1, min_row=16, max_row=25)
    chart.set_categories(cats)

    s3 = chart.series[0]
    s3.graphicalProperties.line.solidFill = "cfcfea"  # light blue
    s8 = chart.series[1]
    s8.graphicalProperties.line.solidFill = "5097a4"  # medium blue
    s9 = chart.series[2]
    s9.graphicalProperties.line.solidFill = "0e2f44"  # dark blue'''

    ws.add_chart(chart, "H15")

    '''process for creating income chart'''

    baseline_total_income = calculate_income_totals(name, baseline_fin_data)
    last_q_total_income = calculate_income_totals(name, last_fin_data)
    latest_q_total_income = calculate_income_totals(name, latest_fin_data)

    total_income_list = [baseline_total_income, last_q_total_income, latest_q_total_income]

    if sum(latest_q_total_income) is not 0:
        for x, l in enumerate(total_income_list):
            for i, total in enumerate(l):
                ws.cell(row=i + 32, column=x + 2, value=total)

        '''data for graph labeling'''

        for i, quarter in enumerate(labeling_list_quarter):
            ws.cell(row=32, column=i + 2, value=quarter)

        for i, label in enumerate(labeling_list_year):
            ws.cell(row=32 + i, column=1, value=label)


        '''income graph'''

        chart = LineChart()
        chart.title = str(name) + ' Income Profile'
        chart.style = 4
        chart.x_axis.title = 'Financial Year'
        chart.y_axis.title = 'Cost £m'

        font = Font(typeface='Calibri')
        size = 1200  # 12 point size
        cp = CharacterProperties(latin=font, sz=size, b=True)  # Bold
        pp = ParagraphProperties(defRPr=cp)
        rtp = RichText(p=[Paragraph(pPr=pp, endParaRPr=cp)])
        chart.x_axis.title.tx.rich.p[0].pPr = pp
        chart.y_axis.title.tx.rich.p[0].pPr = pp
        # chart.title.tx.rich.p[0].pPr = pp

        # title
        size_2 = 1400
        cp_2 = CharacterProperties(latin=font, sz=size_2, b=True)
        pp_2 = ParagraphProperties(defRPr=cp_2)
        rtp_2 = RichText(p=[Paragraph(pPr=pp_2, endParaRPr=cp_2)])
        chart.title.tx.rich.p[0].pPr = pp_2

        #unprofiled costs not included in the chart
        data = Reference(ws, min_col=2, min_row=32, max_col=4, max_row=42)
        chart.add_data(data, titles_from_data=True)
        cats = Reference(ws, min_col=1, min_row=33, max_row=42)
        chart.set_categories(cats)

        
        '''
        keeping as colour coding is useful
        s1 = chart.series[0]
        s1.graphicalProperties.line.solidFill = "cfcfea" #light blue
        s2 = chart.series[1]
        s2.graphicalProperties.line.solidFill = "e2f1bb" #light green 
        s3 = chart.series[2]
        s3.graphicalProperties.line.solidFill = "eaba9d" #light red
        s4 = chart.series[3]
        s4.graphicalProperties.line.solidFil = "5097a4" #medium blue
        s5 = chart.series[4]
        s5.graphicalProperties.line.solidFill = "a0db8e" #medium green
        s6 = chart.series[5]
        s6.graphicalProperties.line.solidFill = "b77575" #medium red
        s7 = chart.series[6]
        s7.graphicalProperties.line.solidFil = "0e2f44" #dark blue
        s8 = chart.series[7]
        s8.graphicalProperties.line.solidFill = "29ab87" #dark green
        s9 = chart.series[8]
        s9.graphicalProperties.line.solidFill = "691c1c" #dark red
        '''

        s3 = chart.series[0]
        s3.graphicalProperties.line.solidFill = "e2f1bb"  # light green
        s8 = chart.series[1]
        s8.graphicalProperties.line.solidFill = "a0db8e"  # medium green
        s9 = chart.series[2]
        s9.graphicalProperties.line.solidFill = "29ab87"  # dark green

        ws.add_chart(chart, "H31")

    else:
        pass

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


# TODO add income


''' ONE: master data to be used for analysis'''

latest_q_data = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\master_3_2018.xlsx')
last_q_data = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\master_2_2018.xlsx')
yearago_q_data = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\master_3_2017.xlsx')

'''TWO: project name list options - this is where the group of interest is specified '''

'''option 1 - all '''
proj_names_all = list(latest_q_data.keys())

'''option 2 - a group'''
#TODO write function for filtering list of project names based on group
#proj_names_group

'''option 3 - bespoke list of projects'''
proj_names_bespoke = ['']

'''THREE: enter variables created via options above into functions and run programme'''

latest_financial_data = financial_dict(proj_names_all, latest_q_data, all_data_lists)
last_financial_data = financial_dict(proj_names_all, last_q_data, all_data_lists)
yearago_financial_data = financial_dict(proj_names_all, yearago_q_data, all_data_lists)

'''FOUR: run the programme'''

for project in proj_names_all:
    wb = place_in_excel(project, latest_financial_data, last_financial_data, yearago_financial_data)
    wb.save('C:\\Users\\Standalone\\Will\\Q3_1819_{}_financial profile.xlsx'.format(project))