import os
from create_wp_json.parser_csv import ParserSCV


def make_json_org(name):
    print(name)
    output_dir = f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\{name}'
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    csv_address_str = f'input/{name}.csv'
    parser_csv = ParserSCV(csv_address_str)
    # Обязательные параметры:
    # По строкам 'row', по столбцам 'column'
    flag = 'column'
    #   Введите количество колонок с подразделениями
    count = 1
    #   Введите порядковый номер колонки с названиями рабочих мест
    rm_column = 1
    #   Введите номер колонки с просчетом
    count_column = 2

    # Не обязательные колонки, если колонки нет введите 0:
    #   Введите порядковый номер колонки с ФИО.
    fio_column = 0
    #   Введите порядковый номер колонки с СНИЛС.
    snils_column = 0
    #   Введите порядковый номер колонки с оборудованием.
    oborud_column = 4
    #   Введите порядковый номер колонки с материалами
    material_column = 5
    #   Введите порядковый номер колонки с индивидуальным номером
    ind_code_column = 0
    #   Введите порядковый номер колонки с типом рабочего места (должно соответствовать ключу в dict.json)
    rm_type_column = 7
    #   Введите порядковый номер колонки с кодом по ОК
    codeok_column = 0
    #   Введите порядковый номер колонки с ЕТКС
    etks_column = 0
    #   Введите порядковый номер колонки с адресом
    address_column = 0
    #   Введите порядковый номер колонки с временем смены
    timesmena_column = 0
    #   Введите порядковый номер колонки со сменностью
    people_in_rm_column = 3

    if flag == 'column':
        # Вызов парсера по столбцам
        parser_csv.column_parsing(count, rm_column, count_column, fio_column=fio_column, snils_column=snils_column,
                                  oborud_column=oborud_column, material_column=material_column,
                                  ind_code_column=ind_code_column, rm_type_column=rm_type_column,
                                  etks_column=etks_column, codeok_column=codeok_column, address_column=address_column,
                                  timesmena_column=timesmena_column, is_address_in_dep=True,
                                  people_in_rm_column=people_in_rm_column)
    elif flag == 'row':
        # Вызов парсера по строкам
        parser_csv.row_parsing(rm_column, count_column, fio_column=fio_column, snils_column=snils_column,
                               oborud_column=oborud_column, material_column=material_column,
                               ind_code_column=ind_code_column, rm_type_column=rm_type_column,
                               etks_column=etks_column, codeok_column=codeok_column, address_column=address_column,
                               timesmena_column=timesmena_column, people_in_rm_column=people_in_rm_column)

    parser_csv.get_json(output_dir + '\\dict.json')
