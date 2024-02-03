from constants import RANKS
from utils import get_parmasters_count, get_data, get_files
import re


class Parmaster:
    """Класс, представляющий Parmaster."""

    def __init__(self, name, rank, author_procedures, collective_procedures):
        self.name = name
        self.rank = rank
        self.author_procedures = author_procedures
        self.collective_procedures = collective_procedures
        self.base_salary = RANKS[self.rank]['base_salary']
        self.calculated_stake = self.calculate_stake()[0]
        self.shifts = self.count_shifts()
        self.procedure_percentage = RANKS[self.rank]['procedure_percentage']
        self.percentage_for_author = 0

    def count_shifts(self):
        """Подсчитать количество смен."""
        return self.calculate_stake()[1] + \
            self.calculate_author_procedures()[0][0] + self.calculate_author_procedures()[1][0] + \
            self.calculate_collective_procedures()[0]

    def calculate_stake(self):
        """Получение прибавки."""
        parmasters_count = get_parmasters_count(get_data(get_files('tmp*.xlsx')[0]))
        stake = 0
        counter = 0
        for name_to_find in list(parmasters_count.keys()):
            name_parts_to_find = re.split(r'\s+', name_to_find.strip())
            name_parts = re.split(r'\s+', self.name.lower())
            for part in name_parts_to_find:
                if part.lower() in name_parts:
                    stake = RANKS[self.rank]['base_salary'] * parmasters_count[name_to_find]
                    counter = parmasters_count[name_to_find]
                    break
        return stake, counter

    def calculate_salary(self):
        """Рассчитать зарплату Parmaster."""
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
        total_salary = ((
                                author_earnings + collective_earnings) * self.procedure_percentage if self.procedure_percentage != 0 else (
                author_earnings + collective_earnings)) + self.calculated_stake + self.calculate_author_percentage() # Если стажерам накидывается 10%
        # total_salary = ((
        #                         author_earnings + collective_earnings) * self.procedure_percentage if self.procedure_percentage != 0 else (
        #         author_earnings + collective_earnings)) + self.calculated_stake + self.calculate_author_percentage() # Если стажерам не накидывается 10%
        total_salary_no_percent = author_earnings + collective_earnings + self.calculated_stake
        return total_salary, total_salary_no_percent

    def calculate_author_percentage(self):
        """Рассчитать процент НЕ автора."""
        author_percentage = 0
        for procedure_type, procedures in self.author_procedures.items():
            for name, count, salary_value in procedures:
                if name == self.name:
                    author_percentage = 0
                    break
                if procedure_type == 'Парение авторское' and name != self.name:
                    author_percentage += salary_value * 0.1
                elif procedure_type == 'Коллективное парение для компании' and name != self.name:
                    author_percentage += salary_value * 0.1
        return author_percentage

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

    def calculate_author_procedures(self):
        """Рассчитать количество и зарплату за авторские процедуры."""
        author_count = 0
        author_collective_count = 0
        author_earnings = 0
        author_collective_earnings = 0
        for procedure_type, procedures in self.author_procedures.items():
            for name, count, salary_value in procedures:
                if procedure_type == 'Парение авторское' and name == self.name:
                    author_count = count
                    author_earnings = salary_value * 0.3
                elif procedure_type == 'Коллективное парение для компании' and name == self.name:
                    author_collective_count = count
                    author_collective_earnings = salary_value * 0.3
        return (author_count, author_earnings), (author_collective_count, author_collective_earnings)

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
