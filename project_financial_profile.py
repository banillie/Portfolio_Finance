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
from aggregate_financial_profile import financial_info

# TODO change/tweak how it is designed so can use functions from aggregate_financial_profile


def place_in_excel(list_names, cells_to_capture, latest_fin_data, last_fin_data, baseline_fin_data):
    wb = Workbook()
    ws = wb.active
    row_start = 3

    '''
    function places all data into excel spreadsheet and creates chart.
    data is placed into sheet in reverse order so that most recent
    data is displayed on right hand side of the data table
    '''

    for name in list_names:
        for key in cells_to_capture:
            ws.cell(row=row_start, column=11, value=latest_financial_data[name][key])
            ws.cell(row=row_start, column=12, value=two[x][1])
            ws.cell(row=row_start, column=13, value=three[x][1])
            ws.cell(row=row_start, column=14, value=(one[x][1] + two[x][1] + three[x][1]))
        try:
            ws.cell(row=row_start, column=7, value=four[x][1])
        except IndexError:
            ws.cell(row=row_start, column=7, value=0)
        try:
            ws.cell(row=row_start, column=8, value=five[x][1])
        except IndexError:
            ws.cell(row=row_start, column=8, value=0)
        try:
            ws.cell(row=row_start, column=9, value=six[x][1])
        except IndexError:
            ws.cell(row=row_start, column=9, value=0)
        try:
            ws.cell(row=row_start, column=10, value=(four[x][1] + five[x][1] + six[x][1]))
        except IndexError:
            ws.cell(row=row_start, column=10, value=0)
        try:
            ws.cell(row=row_start, column=3, value=seven[x][1])
        except IndexError:
            ws.cell(row=row_start, column=3, value=0)
        try:
            ws.cell(row=row_start, column=4, value=eight[x][1])
        except IndexError:
            ws.cell(row=row_start, column=4, value=0)
        try:
            ws.cell(row=row_start, column=5, value=nine[x][1])
        except IndexError:
            ws.cell(row=row_start, column=5, value=0)
        try:
            ws.cell(row=row_start, column=6, value=(seven[x][1] + eight[x][1] + nine[x][1]))
        except IndexError:
            ws.cell(row=row_start, column=6, value=0)
        row_start += 1

    ws.cell(row=2, column=3, value='RDEL')
    ws.cell(row=2, column=4, value='CDEL')
    ws.cell(row=2, column=5, value='Non-Gov')
    ws.cell(row=2, column=6, value='Profile - one year ago')
    ws.cell(row=2, column=7, value='RDEL')
    ws.cell(row=2, column=8, value='CDEL')
    ws.cell(row=2, column=9, value='Non-Gov')
    ws.cell(row=2, column=10, value='Profile - last quarter')
    ws.cell(row=2, column=11, value='RDEL')
    ws.cell(row=2, column=12, value='CDEL')
    ws.cell(row=2, column=13, value='Non-Gov')
    ws.cell(row=2, column=14, value='Profile - current')

    ws.cell(row=2, column=2, value='Spend')
    # ws.cell(row=3, column=2, value='17/18')
    ws.cell(row=3, column=2, value='18/19')
    ws.cell(row=4, column=2, value='19/20')
    ws.cell(row=5, column=2, value='20/21')
    ws.cell(row=6, column=2, value='21/22')
    ws.cell(row=7, column=2, value='22/23')
    ws.cell(row=8, column=2, value='23/24')
    ws.cell(row=9, column=2, value='24/25')
    ws.cell(row=10, column=2, value='25/26')
    ws.cell(row=11, column=2, value='26/27')
    ws.cell(row=12, column=2, value='27/28')
    ws.cell(row=13, column=2, value='Unprofiled')

    ws.cell(row=1, column=3, value='One year ago')
    ws.cell(row=1, column=7, value='Laster quarter')
    ws.cell(row=1, column=11, value='Current quarter')

    ws.merge_cells(start_row=1, start_column=3, end_row=1, end_column=6)
    ws.merge_cells(start_row=1, start_column=7, end_row=1, end_column=10)
    ws.merge_cells(start_row=1, start_column=11, end_row=1, end_column=14)

    '''process for showing total cost profile - and then graph created'''

    row_start = 16

    for x in range(0, len(one)):
        ws.cell(row=row_start, column=5, value=(one[x][1] + two[x][1] + three[x][1]))
        try:
            ws.cell(row=row_start, column=4, value=(four[x][1] + five[x][1] + six[x][1]))
        except IndexError:
            ws.cell(row=row_start, column=4, value=0)
        try:
            ws.cell(row=row_start, column=3, value=(seven[x][1] + eight[x][1] + nine[x][1]))
        except IndexError:
            ws.cell(row=row_start, column=3, value=0)
        row_start += 1

    ws.cell(row=15, column=3, value='One year ago')
    ws.cell(row=15, column=4, value='Last quarter')
    ws.cell(row=15, column=5, value='Latest')

    ws.cell(row=15, column=2, value='Spend')
    # ws.cell(row=3, column=2, value='17/18')
    ws.cell(row=16, column=2, value='18/19')
    ws.cell(row=17, column=2, value='19/20')
    ws.cell(row=18, column=2, value='20/21')
    ws.cell(row=19, column=2, value='21/22')
    ws.cell(row=20, column=2, value='22/23')
    ws.cell(row=21, column=2, value='23/24')
    ws.cell(row=22, column=2, value='24/25')
    ws.cell(row=23, column=2, value='25/26')
    ws.cell(row=24, column=2, value='26/27')
    ws.cell(row=25, column=2, value='27/28')
    ws.cell(row=26, column=2, value='Unprofiled')

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
    data = Reference(ws, min_col=3, min_row=15, max_col=5, max_row=25)
    chart.add_data(data, titles_from_data=True)
    cats = Reference(ws, min_col=2, min_row=16, max_row=25)
    chart.set_categories(cats)

    s3 = chart.series[0]
    s3.graphicalProperties.line.solidFill = "cfcfea"  # light blue
    s8 = chart.series[1]
    s8.graphicalProperties.line.solidFill = "5097a4"  # medium blue
    s9 = chart.series[2]
    s9.graphicalProperties.line.solidFill = "0e2f44"  # dark blue'''

    ws.add_chart(chart, "H15")

    '''process for creating income chart'''

    '''If statement used to create income charts for only those projects
    reporting income'''
    tally = []
    for i in range(0, len(income_one)):
        tally.append(income_one[i][1])

    row_start = 32

    if sum(tally) != 0:
        for x in range(0, len(one)):
            ws.cell(row=row_start, column=5, value=(income_one[x][1]))
            try:
                ws.cell(row=row_start, column=4, value=(income_two[x][1]))
            except IndexError:
                ws.cell(row=row_start, column=4, value=0)
            try:
                ws.cell(row=row_start, column=3, value=(income_three[x][1]))
            except IndexError:
                ws.cell(row=row_start, column=3, value=0)
            row_start += 1

        ws.cell(row=31, column=3, value='One year ago')
        ws.cell(row=31, column=4, value='Last quarter')
        ws.cell(row=31, column=5, value='Current')

        ws.cell(row=31, column=2, value='Spend')
        # ws.cell(row=3, column=2, value='17/18')
        ws.cell(row=32, column=2, value='18/19')
        ws.cell(row=33, column=2, value='19/20')
        ws.cell(row=34, column=2, value='20/21')
        ws.cell(row=35, column=2, value='21/22')
        ws.cell(row=36, column=2, value='22/23')
        ws.cell(row=37, column=2, value='23/24')
        ws.cell(row=38, column=2, value='24/25')
        ws.cell(row=39, column=2, value='25/26')
        ws.cell(row=40, column=2, value='26/27')
        ws.cell(row=41, column=2, value='27/28')
        ws.cell(row=42, column=2, value='Unprofiled')

        chart = LineChart()
        chart.title = str(name) + ' Income Profile'
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
        data = Reference(ws, min_col=3, min_row=31, max_col=5, max_row=41)
        chart.add_data(data, titles_from_data=True)
        cats = Reference(ws, min_col=2, min_row=32, max_row=41)
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

# financial_year = ['17-18 Forecast Non-Gov', '17-18 CDEL Forecast Total', '17-18 RDEL Forecast Total']

# TODO add income

#real_or_nominal = 'Real or Nominal - Actual/Forecast'

#real_details = ['Index Year', 'Deflator']

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
#proj_names_bespoke = ['Digital Railway']

'''THREE: enter variables created via options above into functions and run programme'''

latest_financial_data = financial_info(proj_names_all, latest_q_data, all_data_lists)
last_financial_data = financial_info(proj_names_all, last_q_data, all_data_lists)
yearago_financial_data = financial_info(proj_names_all, yearago_q_data, all_data_lists)

'''loops through all projects producing indvidual financial profiles.
first produces profile as reported by project, secondly creates real
amended to nominal profiles for those projects reporting real. Documents
are saved accordingly.
for x in project_names:
    print(x)
    one_rdel = financial_info(current_Q_dict[x], capture_rdel)
    one_cdel = financial_info(current_Q_dict[x], capture_cdel)
    one_ng = financial_info(current_Q_dict[x], capture_ng)
    one_income = financial_info(current_Q_dict[x], capture_income)

    try:
        two_rdel = financial_info(last_Q_dict[x], capture_rdel)

    except KeyError:
        two_rdel = []

    try:
        two_cdel = financial_info(last_Q_dict[x], capture_cdel)
    except KeyError:
        two_cdel = []

    try:
        two_ng = financial_info(last_Q_dict[x], capture_ng)
    except KeyError:
        two_ng = []

    try:
        two_income = financial_info(last_Q_dict[x], capture_income)
    except KeyError:
        two_income = []

    try:
        three_rdel = financial_info(yearago_Q_dict[x], capture_rdel)
    except KeyError:
        three_rdel = []

    try:
        three_cdel = financial_info(yearago_Q_dict[x], capture_cdel)
    except KeyError:
        three_cdel = []

    try:
        three_ng = financial_info(yearago_Q_dict[x], capture_ng)
    except KeyError:
        three_ng = []

    try:
        three_income = financial_info(yearago_Q_dict[x], capture_income)
    except KeyError:
        three_income = []

    wb = place_in_excel(x, one_rdel, one_cdel, one_ng, one_income, two_rdel, two_cdel, two_ng, two_income, three_rdel,
                        three_cdel, three_ng, three_income)

    wb.save('C:\\Users\\Standalone\\Will\\Q3_1819_{}_financials.xlsx'.format(x))

    if x in projs_rpting_real:
        print('coverting ' + x + ' figures to nominal')
        one_rdel = real_conversion(x, one_rdel, real_meta)
        one_cdel = real_conversion(x, one_cdel, real_meta)
        one_ng = real_conversion(x, one_ng, real_meta)
        one_income = real_conversion(x, one_income, real_meta)

        try:
            two_rdel = real_conversion(x, two_rdel, real_meta)
        except KeyError:
            two_rdel = []

        try:
            two_cdel = real_conversion(x, two_cdel, real_meta)
        except KeyError:
            two_cdel = []

        try:
            two_ng = real_conversion(x, two_ng, real_meta)
        except KeyError:
            two_ng = []

        try:
            two_income = real_conversion(x, two_income, real_meta)
        except KeyError:
            two_income = []

        try:
            three_rdel = real_conversion(x, three_rdel, real_meta)
        except KeyError:
            three_rdel = []

        try:
            three_cdel = real_conversion(x, three_cdel, real_meta)
        except KeyError:
            three_cdel = []

        try:
            three_ng = real_conversion(x, three_ng, real_meta)
        except KeyError:
            three_ng = []

        try:
            three_income = real_conversion(x, three_income, real_meta)
        except KeyError:
            three_income = []

        wb = place_in_excel(x, one_rdel, one_cdel, one_ng, one_income, two_rdel, two_cdel, two_ng, two_income,
                            three_rdel, three_cdel, three_ng, three_income)

        wb.save('C:\\Users\\Standalone\\Will\\Q3_1819_{}_financials_real_con_nominal.xlsx'.format(x))'''