'''

functions for coverting figures from real to nominal and vis versa. using this as a place to store/dump code for now.
not a working programme

'''

'''function for converting real figures into nominal. the value return i.e. a
is the nominal figures for specified year'''

def convert_real(cost, rate, length):
    number_of_years = []
    year = 1
    '''option to apply or not apply inflation to first year'''
    figure = cost
    # figure = round(cost*rate, 1)

    while year < length:
        figure = round(figure * rate, 1)
        number_of_years.append(figure)
        year += 1

    if len(number_of_years) != 0:
        a = number_of_years[-1]
    else:
        a = figure

    return a

'''function that uses convert_real function above to place coverted figures
firstly into a list'''

def real_conversion(project_name, profile, meta_information):
    index_year = meta_information[project_name][0][1]
    # print(index_year)
    length = 2017 - index_year  # calcuates the length of time to apply inflation
    deflator = meta_information[project_name][1][1] + 1  # one is added to become inflator

    r_list = []
    for i in range(0, len(profile)):
        z = convert_real(profile[i][1], deflator, i + length)
        r_list.append(z)

    another_list = []
    for i in range(0, len(profile)):
        a = r_list[i]
        b = profile[i][0]
        c = (b, a)
        another_list.append(c)

    return another_list

def financial_cat(project_list, data, cell_of_interest, f_type):
    output_list = []
    for x in project_list:
        a = data[x][cell_of_interest]
        if a == f_type:
            output_list.append(x)
    return output_list

'''get lists of which projects are reporting nominal, real, or none'''
projs_rpting_nominal = financial_cat(project_names, current_Q_dict, real_or_nominal, 'Nominal')
projs_rpting_real = financial_cat(project_names, current_Q_dict, real_or_nominal, 'Real')
projs_rpting_tbc = financial_cat(project_names, current_Q_dict, real_or_nominal, None)


'''get the meta_data used to calculate nominal from real'''
real_meta = financial_info_dict(projs_rpting_real, current_Q_dict, real_details)