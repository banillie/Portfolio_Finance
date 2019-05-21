from bcompiler.utils import project_data_from_master
#from openpyxl import Workbook
#from openpyxl.styles import Font
import random


def get_totals(m_dict, rag_list):
    output_list = []

    proj_name = random.choice(list(m_dict.keys()))
    quarter_stamp = m_dict[proj_name]['Reporting period (GMPP - Snapshot Date)']
    output_list.append(quarter_stamp)

    for rag in rag_list:
        total = 0
        for proj_name in m_dict.keys():
            if proj_name in remove_projects:
                pass
            else:
                proj_rag = m_dict[proj_name]['Departmental DCA']
                if proj_rag == rag:
                    proj_total = m_dict[proj_name]['Total Forecast']
                    total = total + proj_total
                else:
                    pass
        output_list.append((rag, total))

    return output_list

def check_rag(m_dict, rag_list):

    for proj_name in m_dict.keys():
        rag = m_dict[proj_name]['Departmental DCA']
        if rag in rag_list:
            pass
        else:
            print(proj_name, rag)


rag_list_dca = ['Red', 'Amber/Red', 'Amber', 'Amber/Green', 'Green']

remove_projects = ['West Coast Partnership Franchise', 'South Eastern Rail Franchise Competition',
                   'Rail Franchising Programme', 'East Midlands Franchise', 'HS2 Phase 2b',
                   'HS2 Phase1', 'HS2 Phase2a']

latest_q = project_data_from_master("C:\\Users\\Standalone\\Will\\masters folder\\core data\\master_4_2018.xlsx")
last_q = project_data_from_master("C:\\Users\\Standalone\\Will\\masters folder\\core data\\master_3_2018.xlsx")
#q2_1819_q= project_data_from_master("C:\\Users\\Standalone\\Will\\masters folder\\core data\\master_2_2018.xlsx")

latest_rag_totals = get_totals(latest_q, rag_list_dca)
last_rag_totals = get_totals(last_q, rag_list_dca)
#q2_rag_totals = get_totals(q2_1819_q, rag_list_dca)

#latest_rag_check = check_rag(latest_q, rag_list_dca)

print(latest_rag_totals)
print(last_rag_totals)