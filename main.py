import os
import pandas as pd
import fnmatch
import inquirer
import traceback
import re
from pathlib import Path


class Parmaster:
    """Класс, представляющий Parmaster."""

    RANKS = {
        'стажер': {'base_salary': 2500, 'procedure_percentage': 0},
        'младший мастер': {'base_salary': 1300, 'procedure_percentage': 0.25},
        'мастер': {'base_salary': 1500, 'procedure_percentage': 0.30},
        'старший мастер': {'base_salary': 2000, 'procedure_percentage': 0.35},
    }

    def __init__(self, name, rank, author_procedures, collective_procedures):
        self.name = name
        self.rank = rank
        self.author_procedures = author_procedures
        self.collective_procedures = collective_procedures

    def calculate_salary(self):
        """Рассчитать зарплату Parmaster."""
        base_salary = self.RANKS[self.rank]['base_salary']
        procedure_percentage = self.RANKS[self.rank]['procedure_percentage']

        author_earnings = sum(
            value
            for procedure_type, procedures in self.author_procedures.items()
            for author_info in procedures
            if author_info[0] == self.name
            for value in author_info[2:]
        )
        collective_earnings = 0
        for procedure_type, procedures in self.collective_procedures.items():
            if not isinstance(procedure_type, str):
                continue
            for name, count in procedures:
                if name == self.name:
                    if 'Русская' in procedure_type:
                        collective_earnings += count * 600
                    else:
                        collective_earnings += count * 400
        total_salary = author_earnings + collective_earnings

        return int(total_salary), base_salary

    def calculate_author_procedures(self):
        """Рассчитать количество и зарплату за авторские процедуры."""
        author_count = sum(
            count
            for procedure_type, procedures in self.author_procedures.items()
            for name, count, salary_value in procedures
            if name == self.name
        )

        author_earnings = sum(
            salary_value
            for procedure_type, procedures in self.author_procedures.items()
            for name, count, salary_value in procedures
            if name == self.name
        )
        return author_count, author_earnings

    def calculate_collective_procedures(self):
        """Рассчитать количество и зарплату за коллективные процедуры."""
        collective_count = sum(
            count
            for procedure_type, procedures in self.collective_procedures.items()
            for name, count in procedures
            if name == self.name
        )
        collective_earnings = 0
        for procedure_type, procedures in self.collective_procedures.items():
            if not isinstance(procedure_type, str):
                continue
            for name, count in procedures:
                if name == self.name:
                    if 'Русская' in procedure_type:
                        collective_earnings += count * 600
                    else:
                        collective_earnings += count * 400

        return collective_count, collective_earnings

    def calculate_detailed_procedures(self):
        """Рассчитать детализированную информацию о процедурах для вывода в Excel."""
        detailed_procedures = []
        for procedure_type, procedures in self.author_procedures.items():
            for name, count, salary_value in procedures:
                detailed_procedures.append({
                    'Имя': name,
                    'Тип процедуры': procedure_type,
                    'Количество': count,
                    'Зарплата за процедуру': salary_value
                })

        for procedure_type, procedures in self.collective_procedures.items():
            for name, count in procedures:
                if not isinstance(procedure_type, str):
                    continue
                detailed_procedures.append({
                    'Имя': name,
                    'Тип процедуры': procedure_type,
                    'Количество': count,
                    'Зарплата за процедуру': count * 600 if 'Русская' in procedure_type else count * 400
                })

        detailed_procedures_set = set(tuple(sorted(d.items(), key=lambda item: item[0])) for d in detailed_procedures)

        detailed_procedures = [dict(t) for t in detailed_procedures_set]
        return detailed_procedures


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


def get_parmasters(parmasters_info, author_procedures, collective_procedures, debug_mode=False):
    """Получить Parmasters из указанной информации."""
    parmasters = []
    valid_ranks = ['стажер', 'младший мастер', 'мастер', 'старший мастер']
    parmasters_names = parmasters_info[0].iloc[:, 3].dropna().unique()
    for name in parmasters_names:
        if name != 'Продано с блюдом':
            if debug_mode:
                rank = 'мастер'
            else:
                questions = [
                    inquirer.List('rank',
                                  message=f"Выберите ставку сотрудника {name}",
                                  choices=valid_ranks,
                                  ),
                ]
                answers = inquirer.prompt(questions)
                rank = answers['rank']

            parmaster_info_collective = parmasters_info[0][parmasters_info[0].iloc[:, 3] == name]
            for index, row in parmaster_info_collective.iterrows():
                procedure_type = row[2]
                try:
                    count = int(row[4])
                except ValueError:
                    count = 0
                collective_procedures.setdefault(procedure_type, []).append((name, count))
            parmaster = Parmaster(name, rank, author_procedures, collective_procedures)
            parmasters.append(parmaster)

    return parmasters


def get_results(parmasters):
    """Получить результаты от указанных Parmasters."""
    return [{'Имя': parmaster.name, 'Ставка': parmaster.rank,
             'Коллективное парение': parmaster.calculate_collective_procedures()[1],
             'Парение авторское (кол-во)': parmaster.calculate_author_procedures()[0],
             'Парение авторское (сумма)': parmaster.calculate_author_procedures()[1],
             'Котёл': parmaster.calculate_salary()[0],
             'Итого': parmaster.calculate_salary()[1] + parmaster.calculate_salary()[0]} for parmaster in
            parmasters]


def save_results(results, detailed_procedures, file_name):
    """Сохранить результаты в указанный файл."""
    results_df = pd.DataFrame(results)
    detailed_procedures_df = pd.DataFrame(detailed_procedures)

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        results_df.to_excel(writer, sheet_name='Отчет', index=False)
        detailed_procedures_df.to_excel(writer, sheet_name='Все процедуры', index=False)


def main():
    """Основная функция."""

    debug_mode = False  # Включить/выключить режим отладки

    data1_files = get_files('!1*.xlsx')
    data2_files = get_files('!3*.xlsx')
    data3_files = get_files('!4*.xlsx')

    if not data1_files or not data2_files or not data3_files:
        print("Не найден один из файлов")
        user_input = input("Нажмите Enter, чтобы продолжить...")
        exit()
    else:
        data2 = data2_files[0]
        data3 = data3_files[0]

    data2_df = get_data(data2, header=3)
    data3_df = get_data(data3, header=3)

    collective_procedures = get_collective_procedures(data3_df)
    author_procedures = get_author_procedures(data2_df)

    parmasters_info = get_parmasters_info(data3_df, data2_df)
    parmasters = get_parmasters(parmasters_info, author_procedures, collective_procedures, debug_mode)

    results = get_results(parmasters)
    match = re.search(r'\d{2}_\d{2}_\d{4}_\d{2}_\d{2}_\d{2}', data3)
    date_time = match.group() if match is not None else 0
    result_file_name = f'отчет_{date_time}.xlsx'

    detailed_procedures = []
    for parmaster in parmasters:
        detailed_procedures.extend(parmaster.calculate_detailed_procedures())

    save_results(results, detailed_procedures, result_file_name)


if __name__ == "__main__":
    main()
