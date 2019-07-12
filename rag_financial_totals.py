''' Programme that calculates the total portfolio financial value under a specified DCA rating '''


from bcompiler.utils import project_data_from_master
import random


def get_totals(quarter_data_dict, rag_list, rag_of_interest):
    output_list = []

    proj_name = random.choice(list(quarter_data_dict.keys()))
    quarter_stamp = quarter_data_dict[proj_name]['Reporting period (GMPP - Snapshot Date)']
    output_list.append(quarter_stamp)

    for rag in rag_list:
        total = 0
        for proj_name in quarter_data_dict.keys():
            if proj_name in remove_projects:
                pass
            else:
                proj_rag = quarter_data_dict[proj_name][rag_of_interest]
                if proj_rag == rag:
                    proj_total = quarter_data_dict[proj_name]['Total Forecast']
                    total = total + proj_total
                else:
                    pass
        output_list.append((rag, total))

    return output_list

remove_projects = ['West Coast Partnership Franchise', 'South Eastern Rail Franchise Competition',
                   'Rail Franchising Programme', 'East Midlands Franchise', 'HS2 Phase 2b',
                   'HS2 Phase1', 'HS2 Phase2a']

rag_list_five = ['Red', 'Amber/Red', 'Amber', 'Amber/Green', 'Green']
rag_list_three = ['Red', 'Amber', 'Green']

chosen_q_data = project_data_from_master("C:\\Users\\Standalone\\Will\\masters folder\\core data\\master_4_2018.xlsx")

run = get_totals(chosen_q_data, rag_list_three, 'SRO Finance confidence')

print(run)