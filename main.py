import json
import os
import shutil
from datetime import datetime

import pyodbc

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
    count = 3
    #   Введите порядковый номер колонки с названиями рабочих мест
    rm_column = 3
    #   Введите номер колонки с просчетом
    count_column = 11

    # Не обязательные колонки, если колонки нет введите 0:
    #   Введите порядковый номер колонки с ФИО.
    fio_column = 0
    #   Введите порядковый номер колонки с СНИЛС.
    snils_column = 0
    #   Введите порядковый номер колонки с оборудованием.
    oborud_column = 8
    #   Введите порядковый номер колонки с материалами
    material_column = 9
    #   Введите порядковый номер колонки с индивидуальным номером
    ind_code_column = 0
    #   Введите порядковый номер колонки с типом рабочего места (должно соответствовать ключу в dict.json)
    rm_type_column = 10
    #   Введите порядковый номер колонки с кодом по ОК
    codeok_column = 7
    #   Введите порядковый номер колонки с ЕТКС
    etks_column = 6
    #   Введите порядковый номер колонки с адресом
    address_column = 0

    # Вызов парсера
    parser_csv.column_parsing(count, rm_column, count_column, fio_column=fio_column, snils_column=snils_column,
                              oborud_column=oborud_column, material_column=material_column,
                              ind_code_column=ind_code_column, rm_type_column=rm_type_column,
                              etks_column=etks_column, codeok_column=codeok_column, address_column=address_column)
    parser_csv.get_json(output_dir + '\\dict.json')


def select_all_from_DB(cursor, sql_str, *args):
    try:
        row = cursor.execute(sql_str, *args).fetchall()
        if row is not None:
            return row
        else:
            return 0

    except pyodbc.Error as e:
        print("Error in Connection", e)


def create_template_for_template():
    conn_str = f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\input\\template'
    db_conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}}; DBQ={conn_str}\\ARMv51.MDB;'
    res_dict = {}
    json_address = 'C:\\Users\\buhu_\\PycharmProjects\\OPR\\data\\res_template.json'
    try:
        connect = pyodbc.connect(db_conn_str)
        cursor = connect.cursor()
        # Загружаем файл с таблицами и столбцами:
        sql_dict_json_file = 'C:\\Users\\buhu_\\PycharmProjects\\OPR\\data\\create_template.json'

        with open(sql_dict_json_file, 'r', encoding="utf-8-sig") as file:
            sql_dict = json.load(file)

        for key in sql_dict.keys():
            read_table(cursor, key, sql_dict[key], res_dict)

        try:
            with open(json_address, 'w', encoding="utf-8-sig") as file:
                json.dump(res_dict, file, ensure_ascii=False, indent=4)
        except TypeError as e:
            print(e)

    except pyodbc.Error as e:
        print("Error in Connection", e)


def read_table(cursor, table, sql_str, res_dict):
    rows = cursor.execute(sql_str).fetchall()
    res_dict[table] = []
    for row in rows:
        res_list = []
        for elem in row:
            if type(elem) == datetime:
                res_list.append(str(elem))
            else:
                res_list.append(elem)
        res_dict[table].append(res_list)


if __name__ == '__main__':
    # make_json_org('Винлаб СЗ')
    import_in_db('Винлаб ВЦ')
    # create_template_for_template()
