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

    def __init__(self, name, rank, total_procedures, author_procedures, collective_procedures):
        self.name = name
        self.rank = rank
        self.total_procedures = total_procedures
        self.author_procedures = author_procedures
        self.collective_procedures = collective_procedures

    def calculate_salary(self):
        """Рассчитать зарплату Parmaster."""
        base_salary = self.RANKS[self.rank]['base_salary']
        procedure_percentage = self.RANKS[self.rank]['procedure_percentage']

        procedure_earnings = self.total_procedures * procedure_percentage
        author_earnings = sum(
            value
            for procedure_type, procedures in self.author_procedures.items()
            for author_info in procedures
            if author_info[0] == self.name
            for value in author_info[2:]
        )

        collective_earnings = sum(
            count * 600 if 'Русская' in procedure_type else count * 400
            for procedure_type, procedures in self.collective_procedures.items()
            for name, count in procedures
            if name == self.name
        )

        total_salary = base_salary + procedure_earnings + author_earnings + collective_earnings

        return total_salary

    def calculate_author_procedures(self):
        """Рассчитать количество и зарплату за авторские процедуры."""
        author_count = sum(
            count
            for procedure_type, procedures, value in self.author_procedures.items()
            for name, count in procedures
            if name == self.name
        )
        author_earnings = author_count * self.RANKS[self.rank]['procedure_percentage']
        return author_count, author_earnings

    def calculate_collective_procedures(self):
        """Рассчитать количество и зарплату за коллективные процедуры."""
        collective_count = sum(
            count
            for procedure_type, procedures in self.collective_procedures.items()
            for name, count in procedures
            if name == self.name
        )
        collective_earnings = sum(
            count * 600 if 'Русская' in procedure_type else count * 400
            for procedure_type, procedures in self.collective_procedures.items()
            for name, count in procedures
            if name == self.name
        )
        return collective_count, collective_earnings

    def calculate_detailed_procedures(self):
        """Рассчитать детализированную информацию о процедурах для вывода в Excel."""
        detailed_procedures = []

        for procedure_type, procedures, value in self.author_procedures.items():
            for name, count, salary_value in procedures:
                detailed_procedures.append({
                    'Имя': name,
                    'Тип процедуры': procedure_type,
                    'Количество': count,
                    'Зарплата за процедуру': salary_value
                })

        for procedure_type, procedures in self.collective_procedures.items():
            for name, count in procedures:
                detailed_procedures.append({
                    'Имя': name,
                    'Тип процедуры': procedure_type,
                    'Количество': count,
                    'Зарплата за процедуру': count * 600 if 'Русская' in procedure_type else count * 400
                })

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

    for procedure_type, name, procedure, salary in zip(data.iloc[:, 0], data.iloc[:, 1], data.iloc[:, 2], data.iloc[:, 3]):
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

def get_parmasters_info(data):
    """Получить информацию о Parmasters из указанных данных."""
    return data[data.iloc[:, 3].notna() & data.iloc[:, 4].notna()]

def get_parmasters(parmasters_info, total_procedures, author_procedures, collective_procedures):
    """Получить Parmasters из указанной информации."""
    parmasters = []
    valid_ranks = ['стажер', 'младший мастер', 'мастер', 'старший мастер']
    parmasters_names = parmasters_info.iloc[:, 3].dropna().unique()

    for name in parmasters_names:
        if name != 'Продано с блюдом':
            # questions = [
            #     inquirer.List('rank',
            #                   message=f"Выберите ставку сотрудника {name}",
            #                   choices=valid_ranks,
            #                   ),
            # ]
            # answers = inquirer.prompt(questions)
            # rank = answers['rank']
            rank = 'мастер'
            parmaster_info = parmasters_info[parmasters_info.iloc[:, 1] == name]

            for index, row in parmaster_info.iterrows():
                procedure_type = row[0]
                try:
                    count = int(row[2])
                except ValueError:
                    count = 0
                collective_procedures.setdefault(procedure_type, []).append((name, count))
            parmaster = Parmaster(name, rank, total_procedures, author_procedures, collective_procedures)

            parmasters.append(parmaster)

    return parmasters

def get_results(parmasters):
    """Получить результаты от указанных Parmasters."""
    return [{'Имя': parmaster.name, 'Ставка': parmaster.rank, 'Зарплата': parmaster.calculate_salary()} for parmaster in parmasters]

def save_results(results, file_name):
    """Сохранить результаты в указанный файл."""
    results_df = pd.DataFrame(results)
    results_df.to_excel(file_name, index=False)

def main():
    """Основная функция."""
    data1_files = get_files('!1*.xlsx')
    data2_files = get_files('!3*.xlsx')
    data3_files = get_files('!4*.xlsx')

    if not data1_files or not data2_files or not data3_files:
        print("Не найден один из файлов")
        user_input = input("Нажмите Enter, чтобы продолжить...")
        exit()
    else:
        data1 = data1_files[0]
        data2 = data2_files[0]
        data3 = data3_files[0]

    data1_df = get_data(data1)
    data2_df = get_data(data2)
    data3_df = get_data(data3, header=3)

    collective_procedures = get_collective_procedures(data3_df)
    author_procedures = get_author_procedures(data2_df)


    total_procedures = data1_df.iloc[-1].iloc[2]

    parmasters_info = get_parmasters_info(data3_df)

    parmasters = get_parmasters(parmasters_info, total_procedures, author_procedures, collective_procedures)

    results = get_results(parmasters)

    date_time = re.search(r'\d{2}_\d{2}_\d{4}_\d{2}_\d{2}_\d{2}', data3).group()
    result_file_name = f'отчет_{date_time}.xlsx'

    save_results(results, result_file_name)

if __name__ == "__main__":
    main()