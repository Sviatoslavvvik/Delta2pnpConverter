import csv
import xlrd

# file_format = xlrd.inspect_format(file)
TITLES_COL_POS = {0: "Поз.обозначение",
                  7: "Посадочноеместо",
                  8: "Артикул",
                  1: "Сторона",
                  4: "X'",
                  5: "Y'",
                  6: "Угол"}


def get_data_from_xls(file_name):
    """Получаем данные с листа файла,
    количество строк и количество столбцов"""

    workbook = xlrd.open_workbook_xls(file_name)
    worksheet = workbook.sheet_by_index(0)
    return worksheet, worksheet.ncols, worksheet.nrows


def check_titles(worksheet):
    """Проверяем название столбцов и их порядок"""

    for pos in TITLES_COL_POS.keys():
        temp_str = worksheet.cell_value(0, pos).strip().replace(' ', '')
        if temp_str != TITLES_COL_POS.get(pos):
            raise ValueError(f'{temp_str} не равно {TITLES_COL_POS.get(pos)}')


def check_line(x, y, x_ap, y_ap):
    """Проверяем что координаты X != X' или
    Y != Y' """
    if x != x_ap or y != y_ap:
        return True


def clean_cell(value):
    """Чистим ячейку:
    1.Удаляем лишние пробелы
    2.Удаляем кавычки
    3.заменяем запятые на точки"""
    return value.strip().replace(' ', '').replace('"', '').replace(',', '.')


def get_row_for_write(temp_line):
    """Формируем строку для записи"""
    line_4_save = []
    for ind in TITLES_COL_POS.keys():
        line_4_save.append(clean_cell(temp_line[ind]))
    return line_4_save


def create_file(file_name):
    file = 'input/1.xls'
    worksheet_data, column_count, rows_count = get_data_from_xls(file)
    check_titles(worksheet_data)
    with open('temp.csv', 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter=';')
        csvwriter.writerow(TITLES_COL_POS.values())

        # вставить отступ и реперы
        for line in range(1, rows_count):

            if check_line(
                worksheet_data.cell_value(line, 2),
                worksheet_data.cell_value(line, 3),
                worksheet_data.cell_value(line, 4),
                worksheet_data.cell_value(line, 5),
            ):
                temp_list = worksheet_data.row_values(
                    line, start_colx=0, end_colx=column_count)
                csvwriter.writerow(get_row_for_write(temp_list))
