# Заполнение Расчета, Ведомостей перевески, Акта перевода скота

import openpyxl as opx
# from alive_progress import alive_bar
import pandas as pd
import numpy as np

month = input('Введите отчётный месяц: ')
name = input('Введите Фамилию И.О. животновода: ')
trend = input('Введите направление животных (МКРС/Мясное): ')

df = pd.read_excel(name + ' перевеска.xlsx', 'перевеска')
np_array = df.to_numpy()

df_balance_at_the_beginning_of_the_month  = pd.read_excel(name + ' перевеска.xlsx', 'остаток на начало месяца')
np_array_balance_at_the_beginning_of_the_month = df_balance_at_the_beginning_of_the_month.fillna(0).to_numpy()

df_coming = pd.read_excel(name + ' перевеска.xlsx', 'приход')
np_acoming = df_coming.fillna(0).to_numpy()

df_disposal = pd.read_excel(name + ' перевеска.xlsx', 'выбытие')
np_array_disposal = df_disposal.fillna(0).to_numpy()

df_transference = pd.read_excel(name + ' перевеска.xlsx', 'перевод')
np_array_transference = df_transference.fillna(0).to_numpy()

group_list = np.unique(np_array[:, 2:3])  # получаем уникальные группы животных
# print(group_list)

# bin_count=len(group_list)# получаем количество групп животных
# print(bin_count)

day = '31'
month_num = '01'
year = '22'
last_date = 'на "31" декабря 2021г.'

count_page = 1  # счетчик страниц

wb = opx.load_workbook('sp.xlsx')

sheet_sp_44_1 = wb["44 1"]

for group_s in group_list:

    # фильтр данных по группе
    np_array_filter = np_array[np.in1d(np_array[:, 2], group_s)]
    # print(np_array_filter)

    sheet_sp_43 = wb["43 " + str(count_page)]

    count_line = 1  # счётчик строк
    # all_weight = 0
    # all_gain = 0
    for data_t in np_array_filter:
        # all_weight = all_weight + data_t[3]
        # all_gain = all_gain + data_t[4]

        if count_line < 36:
            sheet_sp_43['E10'] = trend
            sheet_sp_43['E11'] = group_s + ' ' + month
            sheet_sp_43['G12'] = name
            sheet_sp_43['F16'] = name
            sheet_sp_43['P7'] = day
            sheet_sp_43['Q7'] = month_num
            sheet_sp_43['R7'] = year
            sheet_sp_43['F16'], sheet_sp_43['N16'] = last_date, last_date

            sheet_sp_43['A'+str(17 + count_line)] = data_t[1]
            sheet_sp_43['B'+str(17 + count_line)] = count_line
            sheet_sp_43['F'+str(17 + count_line)] = data_t[3]
            sheet_sp_43['G'+str(17 + count_line)] = data_t[4]
            sheet_sp_43['H'+str(17 + count_line)] = data_t[5]
            count_line += 1
        elif 35 < count_line < 71:
            # sheet_sp_43['L'+str(17 + count_line - 35)] = data_t[0]
            sheet_sp_43['L'+str(count_line - 18)] = data_t[1]
            sheet_sp_43['M'+str(count_line - 18)] = count_line
            sheet_sp_43['N'+str(count_line - 18)] = data_t[3]
            sheet_sp_43['O'+str(count_line - 18)] = data_t[4]
            sheet_sp_43['P'+str(count_line - 18)] = data_t[5]
            count_line += 1
        elif 70 < count_line < 106:
            sheet_sp_43['E74'] = trend
            sheet_sp_43['E75'] = group_s + ' ' + month
            sheet_sp_43['G76'] = name
            sheet_sp_43['P71'] = day
            sheet_sp_43['Q71'] = month_num
            sheet_sp_43['R71'] = year
            sheet_sp_43['F80'], sheet_sp_43['N80'] = last_date, last_date

            # sheet_sp_43['A'+str(81 + count_line - 70)] = data_t[0]
            sheet_sp_43['A'+str(11 + count_line)] = data_t[1]
            sheet_sp_43['B'+str(11 + count_line)] = count_line
            sheet_sp_43['F'+str(11 + count_line)] = data_t[3]
            sheet_sp_43['G'+str(11 + count_line)] = data_t[4]
            sheet_sp_43['H'+str(11 + count_line)] = data_t[5]
            count_line += 1
        elif 105 < count_line < 141:
            # sheet_sp_43['L'+str(81 + count_line - 105)] = data_t[0]
            sheet_sp_43['L'+str(count_line - 24)] = data_t[1]
            sheet_sp_43['M'+str(count_line - 24)] = count_line
            sheet_sp_43['N'+str(count_line - 24)] = data_t[3]
            sheet_sp_43['O'+str(count_line - 24)] = data_t[4]
            sheet_sp_43['P'+str(count_line - 24)] = data_t[5]
            count_line += 1
        elif 140 < count_line < 176:
            sheet_sp_43['E135'] = trend
            sheet_sp_43['E136'] = group_s + ' ' + month
            sheet_sp_43['G137'] = name
            sheet_sp_43['P132'] = day
            sheet_sp_43['Q132'] = month_num
            sheet_sp_43['R132'] = year
            sheet_sp_43['F141'], sheet_sp_43['N141'] = last_date, last_date

            # sheet_sp_43['A'+str(142 + count_line - 140)] = data_t[0]
            sheet_sp_43['A'+str(2 + count_line)] = data_t[1]
            sheet_sp_43['B'+str(2 + count_line)] = count_line
            sheet_sp_43['F'+str(2 + count_line)] = data_t[3]
            sheet_sp_43['G'+str(2 + count_line)] = data_t[4]
            sheet_sp_43['H'+str(2 + count_line)] = data_t[5]
            count_line += 1
        elif 175 < count_line < 211:
            # sheet_sp_43['L'+str(142 + count_line - 175)] = data_t[0]
            sheet_sp_43['L'+str(count_line - 33)] = data_t[1]
            sheet_sp_43['M'+str(count_line - 33)] = count_line
            sheet_sp_43['N'+str(count_line - 33)] = data_t[3]
            sheet_sp_43['O'+str(count_line - 33)] = data_t[4]
            sheet_sp_43['P'+str(count_line - 33)] = data_t[5]
            count_line += 1
        elif 210 < count_line < 246:
            sheet_sp_43['E196'] = trend
            sheet_sp_43['E197'] = group_s + ' ' + month
            sheet_sp_43['G198'] = name
            sheet_sp_43['P193'] = day
            sheet_sp_43['Q193'] = month_num
            sheet_sp_43['R193'] = year
            sheet_sp_43['F202'], sheet_sp_43['N202'] = last_date, last_date

            # sheet_sp_43['L'+str(204 + count_line - 210)] = data_t[0]
            sheet_sp_43['L'+str(count_line - 6)] = data_t[1]
            sheet_sp_43['M'+str(count_line - 6)] = count_line
            sheet_sp_43['N'+str(count_line - 6)] = data_t[3]
            sheet_sp_43['O'+str(count_line - 6)] = data_t[4]
            sheet_sp_43['P'+str(count_line - 6)] = data_t[5]
            count_line += 1
        elif 245 < count_line < 281:
            # sheet_sp_43['L'+str(204 + count_line - 245)] = data_t[0]
            sheet_sp_43['L'+str(count_line - 41)] = data_t[1]
            sheet_sp_43['M'+str(count_line - 41)] = count_line
            sheet_sp_43['N'+str(count_line - 41)] = data_t[3]
            sheet_sp_43['O'+str(count_line - 41)] = data_t[4]
            sheet_sp_43['P'+str(count_line - 41)] = data_t[5]
            count_line += 1
            # sheet_sp_43['A'+str(17 + i)] = data_dict.get(i[0])
            # sheet_sp_43['G'+str(17 + i)] = data_dict.get(i[1])

    sum_weight = np.sum(np_array_filter[:, 4:5])
    sum_gain = np.sum(np_array_filter[:, 5:])

    sheet_sp_43['B59'] = sum_weight
    sheet_sp_43['D54'] = sum_gain

    sheet_sp_44_1['M7'], sheet_sp_44_1['R7'] = '"' + day + '" ' + month, year
    sheet_sp_44_1['Z7'], sheet_sp_44_1['AA7'], sheet_sp_44_1['AB7'] = day, month_num, year
    sheet_sp_44_1['C10'], sheet_sp_44_1['D11'], sheet_sp_44_1['A17'] = trend, name, name

    # Остаток на начало месяца
    sheet_sp_44_1['G'+str(count_page + 16)] = group_s
    # голов выбранной группы
    rest_of_heads = int(np_array_balance_at_the_beginning_of_the_month[np.in1d(
        np_array_balance_at_the_beginning_of_the_month[:, 1], group_s)][:, 2:3])
    # живой массы выбранной группы
    remainder_kilogram_1 = int(np_array_balance_at_the_beginning_of_the_month[np.in1d(
        np_array_balance_at_the_beginning_of_the_month[:, 1], group_s)][:, 3:4])
    sheet_sp_44_1['H'+str(count_page + 16)] = rest_of_heads
    sheet_sp_44_1['I'+str(count_page + 16)] = remainder_kilogram_1 

    # Поступление
    # голов выбранной группы в течении месяца
    offspring_heads_1 = int(np.sum(np_acoming[np.in1d(
        np_acoming[:, 2], group_s)][:, 4:5])) # приплод голов в течении месяца
    transference_heads_1 = int(np.sum(np_acoming[np.in1d(
        np_acoming[:, 2], group_s)][:, 8:9])) # переведено голов в течении месяца
    heads_receivedint_1 = offspring_heads_1 + transference_heads_1 
    # живой массы выбранной группы в течении месяца
    offspring_kilogram_1 = int(np.sum(np_acoming[np.in1d(
        np_acoming[:, 2], group_s)][:, 5:6])) # приплод килограмм в течении месяца
    transference_kilogram_1 = int(np.sum(np_acoming[np.in1d(
        np_acoming[:, 2], group_s)][:, 9:10])) # переведено килограмм в течении месяца
    received_kilogram_1 = offspring_kilogram_1 + transference_kilogram_1
    # голов выбранной группы после перевески
    offspring_heads_2 = int(np.sum(np_acoming[np.in1d(
        np_acoming[:, 2], group_s)][:, 6:7])) # приплод голов после перевески
    transference_heads_2 = int(np.sum(np_acoming[np.in1d(
        np_acoming[:, 2], group_s)][:, 10:11])) # переведено голов после перевески
    heads_receivedint_2 = offspring_heads_2 + transference_heads_2
    # живой массы выбранной группы после перевески
    offspring_kilogram_2 = int(np.sum(np_acoming[np.in1d(
        np_acoming[:, 2], group_s)][:, 7:8])) # приплод килограмм после перевески
    transference_kilogram_2 = int(np.sum(np_acoming[np.in1d(
        np_acoming[:, 2], group_s)][:, 11:12])) # переведено килограмм после перевески
    received_kilogram_2 = offspring_kilogram_2 + transference_kilogram_2

    sheet_sp_44_1['J'+str(count_page + 16)
                  ] = heads_receivedint_1 + heads_receivedint_2
    sheet_sp_44_1['K'+str(count_page + 16)
                  ] = received_kilogram_1 + received_kilogram_2

    # Выбытие
    np_array_disposal_filter = np_array_disposal[np.in1d(
        np_array_disposal[:, 2], group_s)]  # фильтр данных по группе выбывших
    
    print(np_array_disposal_filter)

    # ПРОДАННЫХ ГОЛОВ в течении месяца
    heads_sold_1 = int(np.sum(np_array_disposal_filter[:, 4:5]))
    # ПРОДАННЫХ ГОЛОВ после перевески
    heads_sold_2 = int(np.sum(np_array_disposal_filter[:, 6:7]))

    # ПРОДАННОЙ ЖИВОЙ МАССЫ в течении месяца
    sold_kilogram_1 = int(np.sum(np_array_disposal_filter[:, 5:6]))
    # ПРОДАННОЙ ЖИВОЙ МАССЫ после перевески
    sold_kilogram_2 = int(np.sum(np_array_disposal_filter[:, 7:8]))

    # ПЕРЕВЕДЕННЫХ ГОЛОВ в течении месяца
    transferred_heads_1 = int(np.sum(np_array_disposal_filter[:, 8:9]))
    # ПЕРЕВЕДЕННЫХ ГОЛОВ после перевески
    transferred_heads_2 = int(np.sum(np_array_disposal_filter[:, 10:11]))

    # ПЕРЕВЕДЕННОЙ ЖИВОЙ МАССЫ в течении месяца
    transfer_kilogram_1 = int(np.sum(np_array_disposal_filter[:, 9:10]))
    # ПЕРЕВЕДЕННОЙ ЖИВОЙ МАССЫ после перевески
    transfer_kilogram_2 = int(np.sum(np_array_disposal_filter[:, 11:12]))

    # ЗАБИТЫХ ГОЛОВ в течении месяца
    heads_scored_1 = int(np.sum(np_array_disposal_filter[:, 12:13]))
    # ЗАБИТЫХ ГОЛОВ после перевески
    heads_scored_2 = int(np.sum(np_array_disposal_filter[:, 14:15]))

    # ЗАБИТОЙ ЖИВОЙ МАССЫ в течении месяца
    kilogram_scored_1 = int(np.sum(np_array_disposal_filter[:, 13:14]))
    # ЗАБИТОЙ ЖИВОЙ МАССЫ после перевески
    kilogram_scored_2 = int(np.sum(np_array_disposal_filter[:, 15:16]))

    # ПРИРЕЗАННЫХ ГОЛОВ в течении месяца
    heads_slashed_1 = int(np.sum(np_array_disposal_filter[:, 16:17]))
    # ПРИРЕЗАННЫХ ГОЛОВ после перевески
    heads_slashed_2 = int(np.sum(np_array_disposal_filter[:, 18:19]))

    # ПРИРЕЗАННОЙ ЖИВОЙ МАССЫ в течении месяца
    ordered_kilogram_1 = int(np.sum(np_array_disposal_filter[:, 17:18]))
    # ПРИРЕЗАННОЙ ЖИВОЙ МАССЫ после перевески
    ordered_kilogram_2 = int(np.sum(np_array_disposal_filter[:, 19:20]))

    sheet_sp_44_1['L'+str(count_page + 16)] = heads_sold_1 + heads_sold_2 + transferred_heads_1 + \
        transferred_heads_2 + heads_scored_1 + \
        heads_scored_2 + heads_slashed_1 + heads_slashed_2
    out = sold_kilogram_1 + sold_kilogram_2 + transfer_kilogram_1 + \
        transfer_kilogram_2 + kilogram_scored_1 + \
        kilogram_scored_2 + ordered_kilogram_1 + ordered_kilogram_2
    sheet_sp_44_1['M'+str(count_page + 16)] = out
    
    # Выбытие
    
    count_recipient = 1 # счётчик получателя
    group_recipient = np.unique(np_array_transference[:, 4:5])  # получаем уникальных получателей
    
    for recipient in group_recipient:
        count_line = 1  # счётчик строк
        np_array_transference_filter = np_array_transference[np.in1d(np_array_transference[:, 4:5], recipient)]
        df_filter_recipient_group_s = np_array_transference_filter[np.in1d(np_array_transference_filter[:, 2:3], group_s)]
        
        for data_transf in df_filter_recipient_group_s:
            if count_line < 49:
                sheet_sp_47 = wb['47 ' + str(count_recipient) + ' 1']
                sheet_sp_47['C7'] = trend
                sheet_sp_47['B11'] = month
                sheet_sp_47['B12'] = recipient
                sheet_sp_47['B13'] = name
                sheet_sp_47['O8'] = day
                sheet_sp_47['P8'] = month_num
                sheet_sp_47['Q8'] = year
                
                sheet_sp_47['A'+str(18 + count_line)] = data_transf[1]
                sheet_sp_47['B'+str(18 + count_line)] = data_transf[2]
                sheet_sp_47['G'+str(18 + count_line)] = count_line
                sheet_sp_47['H'+str(18 + count_line)] = data_transf[3]
                count_line += 1
            elif 48 < count_line < 101:
                sheet_sp_47 = wb['47 ' + str(count_recipient) + ' 2']
                sheet_sp_47['A'+str(count_line - 48)] = data_transf[1]
                sheet_sp_47['B'+str(count_line - 48)] = data_transf[2]
                sheet_sp_47['J'+str(count_line - 48)] = count_line
                sheet_sp_47['L'+str(count_line - 48)] = data_transf[3]
                #sheet_sp_47['D71'] = day
                #sheet_sp_47['F71'] = month
                #sheet_sp_47['L71'] = year              
                count_line += 1
                
        count_recipient += 1
                  
        

    # Падеж
    # ПАДЕЖ ГОЛОВ в течении месяца
    livestock_mortality_1 = int(np.sum(np_array_disposal_filter[:, 20:21]))
    # ПАДЕЖ ГОЛОВ после перевески
    livestock_mortality_2 = int(np.sum(np_array_disposal_filter[:, 22:23]))

    # ПАДЕЖ ЖИВОЙ МАССЫ в течении месяца
    livestock_death_kilogram_1 = int(
        np.sum(np_array_disposal_filter[:, 21:22]))
    # ПАДЕЖ ЖИВОЙ МАССЫ после перевески
    livestock_death_kilogram_2 = int(np.sum(np_array_disposal_filter[:, 23:]))

    sheet_sp_44_1['N'+str(count_page + 16)
                  ] = livestock_mortality_1 + livestock_mortality_2
    sheet_sp_44_1['O'+str(count_page + 16)
                  ] = livestock_death_kilogram_1 + livestock_death_kilogram_2

    # Остаток на конец месяца
    sheet_sp_44_1['P'+str(count_page + 16)] = rest_of_heads + heads_receivedint_1 + \
        heads_receivedint_2 - heads_sold_1 - heads_sold_2 - transferred_heads_1 - \
        transferred_heads_2 - heads_scored_1 - heads_scored_2 - heads_slashed_1 - \
        heads_slashed_2 - livestock_mortality_1 - livestock_mortality_2
    remainder_kilogram_2 = sum_weight + received_kilogram_2 - sold_kilogram_2 - \
        transfer_kilogram_2 - kilogram_scored_2 - \
        ordered_kilogram_2 - livestock_death_kilogram_2
    sheet_sp_44_1['Q'+str(count_page + 16)] = remainder_kilogram_2
    # Привес
    sheet_sp_44_1['S'+str(count_page + 16)] = remainder_kilogram_2 + livestock_death_kilogram_1 + out - \
        received_kilogram_1 - received_kilogram_2 - remainder_kilogram_1

    filename = 'sp_' + name + ' ' + month_num + year + ' ' + trend + '.xlsx'
    wb.save(filename)

    count_page += 1

wb.close()