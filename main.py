from import_in_access import parsing_csv_to_dic as pars
from import_in_access import import_json

if __name__ == '__main__':
    db_conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};' \
                  r'DBQ=C:\Users\buhu_\PycharmProjects\OPR\DB\ARMv51.MDB;'

    json_address = 'data/res.json'
    con_db = import_json.InterfaceDB(db_conn_str, json_address)
    con_db.insert_ceh(1)

    # csv_address_str = 'data/DL.csv'
    # parser_csv = pars.ParserSCV(csv_address_str)
    # # Обязательные параметры:
    # #   Введите количество колонок с подразделениями
    # count = 0
    # #   Введите порядковый номер колонки с названиями рабочих мест
    # rm_column = 2
    # #   Введите номер колонки с просчетом
    # count_column = 5
    #
    # # Не обязательные колонки, если колонки нет введите 0:
    # #   Введите порядковый номер колонки с ФИО.
    # fio_column = 17
    # #   Введите порядковый номер колонки с СНИЛС.
    # snils_column = 18
    # #   Введите порядковый номер колонки с оборудованием.
    # oborud_column = 15
    # #   Введите порядковый номер колонки с материалами
    # material_column = 16
    # #   Введите порядковый номер колонки с индивидуальным номером
    # ind_code_column = 0
    # #   Введите порядковый номер колонки с типом рабочего места (должно соответствовать ключу в dict.json)
    # rm_type_column = 0
    # #   Введите порядковый номер колонки с кодом по ОК
    # codeok_column = 3
    # #   Введите порядковый номер колонки с ЕТКС
    # etks_column = 4
    #
    #
    # # Вызов парсера
    # parser_csv.row_parsing(rm_column, count_column, fio_column=fio_column, snils_column=snils_column,
    #                        oborud_column=oborud_column, material_column=material_column,
    #                        ind_code_column=ind_code_column, rm_type_column=rm_type_column,
    #                        etks_column=etks_column, codeok_column=codeok_column)
    # parser_csv.get_json('C:\\Users\\buhu_\\PycharmProjects\\OPR\\data\\res.json')

    # # db_conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};' \
    # #               r'DBQ=C:\Users\buhu_\PycharmProjects\OPR\DBksu\ARMv51.MDB;'
    # db_conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};' \
    #               r'DBQ=C:\Users\buhu_\PycharmProjects\OPR\data\DL\ARMv51.MDB;'
    # # csv_address_str = 'First.csv'
    # # organization = 1
    # # csv_address_str = 'Second.csv'
    # # organization = 2
    # # csv_address_str = 'DBksu/ksu.csv'
    # # organization = 1
    # csv_address_str = 'data/DLT.csv'
    # organization = 1
    # # import_in_att(db_conn_str, csv_address_str, organization, 4)
    #
    # import_DL(db_conn_str, csv_address_str, organization)
