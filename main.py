from constants import DEBUG_MODE
from utils import get_files, get_data, get_author_procedures, get_collective_procedures, get_parmasters_info, \
    get_parmasters_count, calculate_cauldron, get_results, save_results
from parmaster import Parmaster
import inquirer
import re
import shutil


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


def main():
    """Основная функция."""

    data1_files = get_files('!1*.xlsx')
    data2_files = get_files('!3*.xlsx')
    data3_files = get_files('!4*.xlsx')
    data4_files = get_files('tmp*.xlsx')

    if not data1_files or not data2_files or not data3_files or not data4_files:
        print("Не найден один из файлов")
        user_input = input("Нажмите Enter, чтобы продолжить...")
        exit()
    else:
        data1 = data1_files[0]
        data2 = data2_files[0]
        data3 = data3_files[0]
        data4 = data4_files[0]

    data1_df = get_data(data1, header=3)
    data2_df = get_data(data2, header=3)
    data3_df = get_data(data3, header=3)
    data4_df = get_data(data4)

    collective_procedures = get_collective_procedures(data3_df)
    author_procedures = get_author_procedures(data2_df)

    parmasters_info = get_parmasters_info(data3_df, data2_df)
    parmasters = get_parmasters(parmasters_info, author_procedures, collective_procedures, DEBUG_MODE)
    parmasters_count = get_parmasters_count(data4_df)

    cauldron = calculate_cauldron(parmasters, parmasters_count, data1_df)
    results = get_results(parmasters, cauldron)

    match = re.search(r'\d{2}_\d{2}_\d{4}_\d{2}_\d{2}_\d{2}', data3)
    date_time = match.group() if match is not None else 'no_date_found'
    result_file_name = f'Финальный_отчет_{date_time}.xlsx'

    detailed_procedures = []
    for parmaster in parmasters:
        detailed_procedures.extend(parmaster.calculate_detailed_procedures())
    save_results(results, detailed_procedures, result_file_name, date_time)
    data_files = data1_files + data2_files + data3_files + data4_files
    for file in data_files:
        shutil.move(file, date_time)


if __name__ == "__main__":
    main()
