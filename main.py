import os
import shutil

from import_in_access import parsing_csv_to_dic as pars
from import_in_access import import_json


def import_in_db(name):
    output_dir = f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\{name}'
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)
    output_dir += '\\DB'
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir, ignore_errors=True)
    shutil.copytree(r'C:\Users\buhu_\PycharmProjects\OPR\data\NewDB', output_dir)
    json_address = f'output\\{name}\\dict.json'

    con_db = import_json.InterfaceDB(output_dir, json_address)
    org_id = con_db.insert_org(name)
    con_db.insert_ceh(org_id)


def make_json_org(name):
    output_dir = f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\{name}'
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    csv_address_str = f'input/{name}.csv'
    parser_csv = pars.ParserSCV(csv_address_str)
    # Обязательные параметры:
    #   Введите количество колонок с подразделениями
    count = 7
    #   Введите порядковый номер колонки с названиями рабочих мест
    rm_column = 7
    #   Введите номер колонки с просчетом
    count_column = 14

    # Не обязательные колонки, если колонки нет введите 0:
    #   Введите порядковый номер колонки с ФИО.
    fio_column = 0
    #   Введите порядковый номер колонки с СНИЛС.
    snils_column = 17
    #   Введите порядковый номер колонки с оборудованием.
    oborud_column = 0
    #   Введите порядковый номер колонки с материалами
    material_column = 0
    #   Введите порядковый номер колонки с индивидуальным номером
    ind_code_column = 16
    #   Введите порядковый номер колонки с типом рабочего места (должно соответствовать ключу в dict.json)
    rm_type_column = 0
    #   Введите порядковый номер колонки с кодом по ОК
    codeok_column = 8
    #   Введите порядковый номер колонки с ЕТКС
    etks_column = 9

    # Вызов парсера
    parser_csv.column_parsing(count, rm_column, count_column, fio_column=fio_column, snils_column=snils_column,
                              oborud_column=oborud_column, material_column=material_column,
                              ind_code_column=ind_code_column, rm_type_column=rm_type_column,
                              etks_column=etks_column, codeok_column=codeok_column)
    parser_csv.get_json(output_dir + '\\dict.json')


if __name__ == '__main__':
    # make_json_org('ПАО ВТБ')
    import_in_db('ПАО ВТБ')
