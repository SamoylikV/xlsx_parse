import os
import pandas as pd
import fnmatch
import re
from collections import Counter
from constants import DEBUG_MODE


def get_files(pattern):
    """Получить файлы, соответствующие заданному шаблону."""
    all_files = os.listdir()
    return [file for file in all_files if fnmatch.fnmatch(file, pattern)]


def get_data(file, header=None):
    """Получить данные из указанного файла."""
    return pd.read_excel(file, header=header)


def get_author_procedures(data):
    """Получить авторские процедуры из указанных данных."""
    author_procedures = {"Коллективное парение для компании": [], "Парение авторское": []}
    current_procedure = ''

    for procedure_type, name, procedure, salary in zip(data.iloc[:, 0], data.iloc[:, 1], data.iloc[:, 2],
                                                       data.iloc[:, 3]):
        if str(procedure).isdigit() is False:
            continue

        if pd.notna(procedure_type):
            current_procedure = procedure_type.strip()

        if pd.notna(name) and pd.notna(procedure):
            procedure_value = int(salary)
            if current_procedure == "Коллективное парение для компании":
                salary_value = round(procedure_value * 0.35, 2)
            elif current_procedure == "Парение авторское":
                salary_value = round(procedure_value * 0.3, 2)
            else:
                salary_value = None

            author_procedures[current_procedure].append((name, procedure, salary_value))

    return author_procedures


def get_collective_procedures(data):
    """Получить коллективные процедуры из указанных данных."""
    collective_procedures = {}
    current_procedure_type = None
    current_names = []
    current_amounts = []

    for proc_type, name, amt in zip(data.iloc[:, 2], data.iloc[:, 3], data.iloc[:, 4]):
        if pd.notna(proc_type):
            collective_procedures[current_procedure_type] = list(zip(current_names, current_amounts))
            if 'Русская' in proc_type or 'Хаммам' in proc_type:
                current_procedure_type = proc_type
                current_names = []
                current_amounts = []
            else:
                current_names.append(name)
                current_amounts.append(amt)
        else:
            if current_procedure_type:
                current_names.append(name)
                current_amounts.append(amt)
    return {key: value for key, value in collective_procedures.items() if value}


def get_parmasters_info(data, data2):
    """Получить информацию о Parmasters из указанных данных."""
    return data[data.iloc[:, 3].notna() & data.iloc[:, 4].notna()], data2.iloc[:, 0], data2.iloc[:, 1], data2.iloc[:,
                                                                                                        2], data2.iloc[
                                                                                                            :, 3]


def get_parmasters_count(data):
    count_notnan = []
    flag = False
    for index, row in data.iterrows():
        if row.iloc[0] == 'Пармастер':
            flag = True
        if flag:
            if row.iloc[0] == 'Системный администратор':
                break
            for col in range(2 + 1, len(data.columns) - 1):
                if pd.notna(row[col]):
                    count_notnan.append(row.iloc[2])

    return dict(Counter(count_notnan))


def calculate_cauldron(parmasters, data1_df):
    """Получение котла."""
    total = data1_df.loc[data1_df.iloc[:, 0] == 'Итого'].iloc[:, 4].iloc[0]
    counter = 0
    # for name_to_find in list(parmasters_count.keys()):
    #     name_parts_to_find = re.split(r'\s+', name_to_find.strip())
    #     for parmaster in parmasters:
    #         name_parts = re.split(r'\s+', parmaster.name.lower())
    #         for part in name_parts_to_find:
    #             if part.lower() in name_parts:
    #                 if parmaster.rank != 'стажер':
    #                     counter += 1
    #                     break
    for _ in parmasters:
        if _.rank != 'стажер':
            counter += _.shifts
    if counter != 0:
        return total / counter
    else:
        return 0


def get_results(parmasters, cauldron):
    """Получить результаты от указанных Parmasters."""
    return [{'Имя': parmaster.name, 'Ставка': parmaster.rank, 'Ставка р.': parmaster.base_salary,
             'Ставка р. (за все смены)': parmaster.calculated_stake,
             '(сумма) Коллективное парение': parmaster.calculate_collective_procedures()[1],
             '(кол-во) Коллективное парение': parmaster.calculate_collective_procedures()[0],
             '(кол-во) Парение авторское': parmaster.calculate_author_procedures()[0][0],
             '(сумма) Парение авторское': parmaster.calculate_author_procedures()[0][1],
             '(кол-во) Парение коллективное авторское': parmaster.calculate_author_procedures()[1][0],
             '(сумма) Парение коллективное авторское': parmaster.calculate_author_procedures()[1][1],
             'Котёл': cauldron * parmaster.procedure_percentage if parmaster.rank != 'стажер' else 0,
             'Итого': parmaster.calculate_salary()[0] + (
                 cauldron * parmaster.procedure_percentage if parmaster.rank != 'стажер' else 0),
             'Кол-во пар': parmaster.shifts + parmaster.calculate_author_procedures()[0][0] +
                            parmaster.calculate_author_procedures()[1][0]
             } for parmaster in parmasters]


def save_results(results, detailed_procedures, file_name, date_time):
    """Сохранить результаты в указанный файл."""
    results_df = pd.DataFrame(results)
    detailed_procedures_df = pd.DataFrame(detailed_procedures)

    if not DEBUG_MODE:
        if not os.path.exists(date_time):
            os.makedirs(date_time)
        file_path = os.path.join(date_time, file_name)
    else:
        file_path = file_name
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        results_df.to_excel(writer, sheet_name='Отчет', index=False)
        detailed_procedures_df.to_excel(writer, sheet_name='Все процедуры', index=False)
