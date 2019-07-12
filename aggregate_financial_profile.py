'''
Programme to create a financial profile for a group of projects i.e. can produce the portfolio profile or a chosen
set of projects profile. It also has the option for comparing like for like projects across projects if necessary

Input documents.
1) Two quarter master data sets

Output documents
2) Excel spreadsheet contain a graph with financial profile

Follow instructions below

'''

from bcompiler.utils import project_data_from_master
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font

def year_totals(project_names_list, cells_to_capture, quarter_data_dict, remove_from_total):
    #TODO convert output into dictionary

    fy_total_list = []

    for key in cells_to_capture:
        thesum = 0
        for name in project_names_list:
            try:
                thesum = thesum + quarter_data_dict[name][key]
            except (TypeError, KeyError):
                pass
        fy_total_list.append(thesum)

    return fy_total_list

def likeforlike():
    '''
    small programme used to filter out projects that are not in both data sets
    :param data_1: most recent quarters data
    :param data_2: less recent quarters data
    :return: a list of projects that are in both data sets
    '''

    one = list(set(latest_q_data) - set(other_q_data))
    two = list(set(other_q_data) - set(latest_q_data))

    output_list = one + two

    return output_list

def place_in_excel(quarter_data_dict, totals_list, cells_to_capture):
    '''
    function places all data into a new workbook.
    :param fin_data: dictionary of project costs created via financial info
    :param totals: list of total costs for each year
    :param cells_to_capture: lists of keys of interest
    :return: excle workbook
    '''

    wb = Workbook()
    ws = wb.active

    ws.cell(row=1, column=1).value = 'Project'
    for i, name in enumerate(quarter_data_dict.keys()):
        '''lists project names in row one'''
        ws.cell(row=1, column=i + 2).value = name

        '''iterates through financial dictionary - placing financial data in ws'''
        for x, key in enumerate(cells_to_capture):
            try:
                ws.cell(row=x+2, column=i+2).value = quarter_data_dict[name][key]
            except KeyError:
                ws.cell(row=x + 2, column=i + 2).value = 0

    '''places totals in final column. to note because this is a list and not a dictionary as for fin_data there is 
    possibility that data could become unaligned. Whether changing the list of cells_to_capture causes them to become
    unaligned needs to be tested'''
    ws.cell(row=1, column=len(quarter_data_dict.keys()) + 2).value = 'Total'
    for i, values in enumerate(totals_list):
        ws.cell(row=i + 2, column=len(quarter_data_dict.keys())+2).value = values

    '''places keys into the chart in the first column'''
    for i, key in enumerate(cells_to_capture):
        ws.cell(row=i+2, column=1).value = key

    '''information on which projects are not included in totals'''
    ws.cell(row=1, column=len(quarter_data_dict.keys()) + 4).value = 'Projects that have been removed to avoid double counting'
    for i, project in enumerate(dont_double_count):
        ws.cell(row=i + 2, column=len(quarter_data_dict.keys()) + 4).value = project

    ws.cell(row=1, column=len(quarter_data_dict.keys())+6).value = 'Projects that have been removed to enable like for like' \
                                                          'comparison of totals'
    for i, project in enumerate(like_for_like_totals):
        ws.cell(row=i + 2, column=len(quarter_data_dict.keys())+6).value = project

    '''data for overall chart. As above because this data is in a list - possibility of it being unaligned needs 
    testing. not the best way of managing data flow, but working for now'''
    start_row = len(totals_list) + 8
    for x in range(0, int(len(totals_list) / 4)):
        ws.cell(row=start_row, column=2, value=totals_list[x])
        start_row += 1

    start_row = len(totals_list) + 8
    for x in range(int(len(totals_list) / 4), (int(len(totals_list) / 4) * 2)):
        ws.cell(row=start_row, column=3, value=totals_list[x])
        start_row += 1

    start_row = len(totals_list) + 8
    for x in range((int(len(totals_list) / 4) * 2), (int(len(totals_list) / 4) * 3)):
        ws.cell(row=start_row, column=4, value=totals_list[x])
        start_row += 1

    start_row = len(totals_list) + 8
    for x in range((int(len(totals_list) / 4) * 3), int(len(totals_list))):
        ws.cell(row=start_row, column=5, value=totals_list[x])
        start_row += 1

    '''code was essentially a hack'''

    start_row = len(totals_list) + 8
    list_of_numbers = [0, len(capture_rdel), len(capture_rdel)*2]
    total_sum = 0
    for i in range(0, len(capture_rdel)):
        for x in list_of_numbers:
            total_sum = total_sum + totals_list[x + i]
            ws.cell(row=start_row, column=6, value=total_sum)
        start_row += 1
        total_sum = 0

    a = len(totals_list) + 7
    ws.cell(row=a, column=2, value='RDEL')
    ws.cell(row=a, column=3, value='CDEL')
    ws.cell(row=a, column=4, value='Non-Gov')
    ws.cell(row=a, column=5, value='Income')
    ws.cell(row=a, column=6, value='Total')


    # ws.cell(row=a+1, column=1, value='17/18')
    #ws.cell(row=a + 1, column=1, value='18/19')
    ws.cell(row=a + 1, column=1, value='19/20')
    ws.cell(row=a + 2, column=1, value='20/21')
    ws.cell(row=a + 3, column=1, value='21/22')
    ws.cell(row=a + 4, column=1, value='22/23')
    ws.cell(row=a + 5, column=1, value='23/24')
    ws.cell(row=a + 6, column=1, value='24/25')
    ws.cell(row=a + 7, column=1, value='25/26')
    ws.cell(row=a + 8, column=1, value='26/27')
    ws.cell(row=a + 9, column=1, value='27/28')
    ws.cell(row=a + 10, column=1, value='28/29')
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

'''List of financial data keys to capture. This should be amended to years of interest'''
capture_rdel = ['19-20 RDEL Forecast Total', '20-21 RDEL Forecast Total', '21-22 RDEL Forecast Total',
                '22-23 RDEL Forecast Total', '23-24 RDEL Forecast Total', '24-25 RDEL Forecast Total',
                '25-26 RDEL Forecast Total', '26-27 RDEL Forecast Total', '27-28 RDEL Forecast Total',
                '28-29 RDEL Forecast Total', 'Unprofiled RDEL Forecast Total']
capture_cdel = ['19-20 CDEL Forecast Total', '20-21 CDEL Forecast Total', '21-22 CDEL Forecast Total',
                 '22-23 CDEL Forecast Total', '23-24 CDEL Forecast Total', '24-25 CDEL Forecast Total',
                 '25-26 CDEL Forecast Total', '26-27 CDEL Forecast Total', '27-28 CDEL Forecast Total',
                 '28-29 CDEL Forecast Total', 'Unprofiled CDEL Forecast Total']
capture_ng = ['19-20 Forecast Non-Gov', '20-21 Forecast Non-Gov', '21-22 Forecast Non-Gov', '22-23 Forecast Non-Gov',
              '23-24 Forecast Non-Gov', '24-25 Forecast Non-Gov', '25-26 Forecast Non-Gov', '26-27 Forecast Non-Gov',
              '27-28 Forecast Non-Gov', '28-29 Forecast Non-Gov', 'Unprofiled Forecast-Gov']
capture_income =['19-20 Forecast - Income both Revenue and Capital',
                '20-21 Forecast - Income both Revenue and Capital', '21-22 Forecast - Income both Revenue and Capital',
                '22-23 Forecast - Income both Revenue and Capital', '23-24 Forecast - Income both Revenue and Capital',
                '24-25 Forecast - Income both Revenue and Capital', '25-26 Forecast - Income both Revenue and Capital',
                '26-27 Forecast - Income both Revenue and Capital', '27-28 Forecast - Income both Revenue and Capital',
                '28-29 Forecast - Income both Revenue and Capital', 'Unprofiled Forecast Income']

all_data_lists = capture_rdel + capture_cdel + capture_ng + capture_income

'''INSTRUCTION FOR RUNNING PROGRAMME'''

'''1) Provide paths to master data sets. Note the second master date set is only used to produce a list of projects
for like for like comparison.'''

latest_q_data = project_data_from_master("C:\\Users\\Standalone\\Will\\masters folder\\core data\\master_4_2018.xlsx")
other_q_data = project_data_from_master("C:\\Users\\Standalone\\Will\\masters folder\\core data\\master_3"
                                 "_2018.xlsx")

''' 2) Choose appropriate project name list options - this is where the group of interest for the aggregate chart is 
specified. '''

'''option 1 - all '''
proj_names_all = list(latest_q_data.keys())

'''option 2 - a group'''
#TODO write function for filtering list of project names based on group
proj_names_group = ['East Midlands Franchise', 'Rail Franchising Programme', 'West Coast Partnership Franchise']

'''option 3 - bespoke list of projects'''
proj_names_bespoke = ['Digital Railway']

'''3) It is important to consider the list of projects that should included within financial totals of each year. There 
are two key things to consider 1) whether project cost profiles should be removed to prevent double counting, 2) 
whether you would like to have a like for like comparison between chosen quarters i.e. compare change in cost profile 
for the same set of projects. 

The options below therefore enable the user to handle these options. '''

'''option one - just remove projects to stop double counting'''
dont_double_count = ['HS2 Phase 2b', 'HS2 Phase1', 'HS2 Phase2a', 'East Midlands Franchise',
                     'South Eastern Rail Franchise Competition', 'West Coast Partnership Franchise']

'''option two - ensure that only like for like comparision of totals.'''
like_for_like_totals = dont_double_count + likeforlike()

'''4) enter variables created via options above into functions and run programme'''

'''step one, run the year_totals function to get year totals. 
variables to be placed in this order: 
(project_names_list, cells_to_capture, quarter_data_dict, remove_from_total)
'''
total_data = year_totals(proj_names_all, all_data_lists, latest_q_data, like_for_like_totals)

'''step two, run the place data_in_excel function to place data in output document.
Variables to be placed in this order:
(quarter_data_dict, totals_list, cells_to_capture)'''
output = place_in_excel(latest_q_data, total_data, all_data_lists)

'''5) specify where to save to output file - excel spreadsheet with graph'''

output.save("C:\\Users\\Standalone\\Will\\Q1_1920_financial_profile_testing.xlsx")