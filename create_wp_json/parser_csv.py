import csv
import json


class ParserSCV:
    def __init__(self, csv_address: str, type='column'):
        self.csv_address = csv_address
        self.res_dict = {'ceh': {}}
        self.type = type

    @staticmethod
    def start_with_b_l(wrong_str: str) -> str:
        w_str = wrong_str.strip()
        return w_str[0].upper() + w_str[1:].strip()
        # .lower()

    def get_json(self, json_address: str):
        with open(json_address, 'w', encoding="utf-8-sig") as file:
            json.dump(self.res_dict, file, ensure_ascii=False, indent=4)

    def row_parsing(self, rm_column: int, count_column: int, fio_column=0, snils_column=0, oborud_column=0,
                    material_column=0, ind_code_column=0, rm_type_column=0, etks_column=0, codeok_column=0,
                    address_column=0, timesmena_column=0, is_address_in_dep=False, people_in_rm_column=0,
                    woman_in_rm_column=0):
        wp_count = 0
        # Открытие файла
        with open(self.csv_address, newline='', encoding="utf-8-sig") as File:
            reader = csv.reader(File, delimiter=';')
            # Задаем начальную точку в словаре
            current = self.res_dict['ceh']
            # Создаем маршрут к текущему отделу в словаре
            current_lvl = 0
            current_list = []
            # Перебираем все строки в документе
            for r in reader:
                # Если в первой колонке есть символ, значит эта строка с отделом.
                if len(r[0]) > 0:
                    # Переводим название отдела в нормальный вид
                    s = self.start_with_b_l(r[1])
                    # Если в первой колонке 0 - это цех
                    if int(r[0]) == 0:
                        # Откат в начало словаря
                        current = self.res_dict['ceh']
                        # Если такого цеха нет, создаем ключ в словаре
                        if s not in current.keys():
                            current[s] = {}
                            current[s]['uch'] = {}
                        # Переводим указатель на текущий цех
                        current = current[s]['uch']
                        # Откатываем маршрут
                        current_lvl = int(r[0])
                        current_list = [s, ]

                    # Если число в первой колонке больше 1 - это участок
                    elif int(r[0]) > 0:
                        # Если уровень текущего участка больше предыдущего — добавляем его
                        if int(r[0]) > current_lvl:
                            if s not in current.keys():
                                current[s] = {}
                            current = current[s]
                            current_lvl = int(r[0])
                            current_list.append(s)

                        # Иначе откатываемся к предыдущему участку с уровнем меньше текущего
                        # И добавляем отдел в него.
                        else:
                            current = self.res_dict['ceh']
                            temp_list = []
                            for i in range(int(r[0])):
                                current = current[current_list[i]]
                                temp_list.append(current_list[i])
                                if i == 0:
                                    current = current['uch']
                            if s not in current.keys():
                                current[s] = {}
                            temp_list.append(s)
                            current = current[s]
                            current_lvl = int(r[0])
                            current_list = temp_list

                else:
                    # Инициализация информации о рабочем месте

                    wp_name = self.start_with_b_l(r[rm_column])
                    wp_count += 1
                    address = f"Фактический адрес: {r[address_column]}" if address_column != 0 \
                        else ''

                    # Вывод адреса в название подразделения, если необходимо
                    # if is_address_in_dep and 'rm' not in current.keys():
                    #     if address not in current.keys():
                    #         current[address] = {}
                    #     current = current[address]
                    #     current_lvl += 1

                    if is_address_in_dep:
                        if address not in current.keys():
                            current[address] = {}
                        current_lvl += 1

                    # Создание ключа рабочих мест, и вставка первого РМ в отдел
                    if 'rm' not in current[address].keys():
                        current[address]['rm'] = []

                    # Если такого рабочего места нет — создаём новую запись
                    # elif len(r[count_column]) > 0:
                    if r[count_column].isdigit():
                        fio = r[fio_column].title() if fio_column != 0 \
                            else ''
                        snils = r[snils_column] if snils_column != 0 \
                            else 'Отсутствует'
                        ind_code = r[ind_code_column] if ind_code_column != 0 \
                            else ''
                        oborud = r[oborud_column] if oborud_column != 0 \
                            else ''
                        material = r[material_column] if material_column != 0 \
                            else ''
                        rm_type = r[rm_type_column].strip() if rm_type_column != 0 \
                            else 'office'
                        etks = r[etks_column] if etks_column != 0 \
                            else ''
                        codeok = r[codeok_column] if codeok_column != 0 \
                            else ''
                        timesmena = int(float(r[timesmena_column].replace(',', '.')) * 60) if timesmena_column != 0 \
                            else 480
                        people_in_rm = int(r[people_in_rm_column]) if people_in_rm_column != 0 else 1
                        # (2 if r[people_in_rm_column] == '2' else
                        #             (4 if r[people_in_rm_column] == '3/3' else 1)) if timesmena_column != 0 \
                        # else 1
                        woman_in_rm = r[woman_in_rm_column] if woman_in_rm_column != 0 \
                            else 0

                        rm_dict = {'caption': wp_name,
                                   'analog': 0,
                                   'oborud': oborud,
                                   'material': material,
                                   'fio': [fio, ],
                                   'snils': [snils, ],
                                   'ind_code': [ind_code, ],
                                   'woman_in_rm': [woman_in_rm, ],
                                   'rm_type': rm_type,
                                   'etks': etks,
                                   'codeok': codeok,
                                   'address': address,
                                   'timesmena': timesmena,
                                   'people_in_rm': people_in_rm}

                        current[address]['rm'].append(rm_dict)

                    # Если место есть — добавляем аналогию
                    else:
                        current[address]['rm'][-1]['analog'] += 1
                        current[address]['rm'][-1]['fio'].append(r[fio_column].title() if fio_column != 0 else '')
                        current[address]['rm'][-1]['snils'].append(r[snils_column] if snils_column != 0
                                                                   else 'Отсутствует')
                        current[address]['rm'][-1]['ind_code'].append(r[ind_code_column] if ind_code_column != 0
                                                                      else '')
                        current[address]['rm'][-1]['woman_in_rm'].append(r[woman_in_rm_column] if
                                                                         woman_in_rm_column != 0 else 0)

        print(f'Создан json файл на {wp_count} рм')
        return wp_count

    def column_parsing(self, count: int, rm_column: int, count_column: int, fio_column=0, snils_column=0,
                       oborud_column=0, material_column=0, ind_code_column=0, rm_type_column=0, etks_column=0,
                       codeok_column=0, address_column=0, timesmena_column=0, is_address_in_dep=False,
                       people_in_rm_column=0, woman_in_rm_column=0):
        wp_count = 0
        # Открытие файла
        with open(self.csv_address, newline='', encoding="utf-8-sig") as File:
            reader = csv.reader(File, delimiter=';')
            # Перебираем все строки в документе
            for r in reader:
                # Задаем начальную точку в словаре
                current = self.res_dict['ceh']

                # Перебираем столбцы с отделами
                for i in range(count):
                    # Если отделы кончились раньше — прерываем цикл
                    if len(r[i]) == 0:
                        break

                    # Переводим название отдела в нормальный вид
                    s = self.start_with_b_l(r[i])

                    # Первая колонка — цех, добавляем его в словарь
                    if i == 0:
                        if s not in current:
                            current[s] = {}
                            current[s]['uch'] = {}
                        current = current[s]['uch']
                    # Остальные колонки — участки, добавляем в словарь
                    else:
                        if s not in current:
                            current[s] = {}
                        current = current[s]

                # Вывод адреса в название подразделения, если необходимо
                if is_address_in_dep:
                    if address_column != 0:
                        address = f"Фактический адрес: {r[address_column]}"
                        if address not in current:
                            current[address] = {}
                        current = current[address]

                if rm_column != 0:
                    wp_count += 1
                    # Инициализация информации о рабочем месте
                    wp_name = self.start_with_b_l(r[rm_column])
                    oborud = self.start_with_b_l(r[oborud_column]) \
                        if (r[oborud_column] != '' and oborud_column != 0)\
                        else ''
                    material = self.start_with_b_l(r[material_column]) \
                        if (r[material_column] != '' and material_column != 0)\
                        else ''
                    fio = r[fio_column] if fio_column != 0 \
                        else ''
                    snils = r[snils_column] if snils_column != 0 \
                        else 'Отсутствует'
                    ind_code = r[ind_code_column] if ind_code_column != 0 \
                        else ''
                    rm_type = r[rm_type_column] if rm_type_column != 0 \
                        else 'office'
                    etks = r[etks_column] if etks_column != 0 \
                        else ''
                    codeok = r[codeok_column] if codeok_column != 0 \
                        else ''
                    address = f"Фактический адрес: {r[address_column]}" if address_column != 0 \
                        else ''
                    timesmena = int(float(r[timesmena_column].replace(',', '.')) * 60) if timesmena_column != 0 \
                        else 480
                    people_in_rm = r[people_in_rm_column] if people_in_rm_column != 0 else 1
                    # people_in_rm = 2 if r[people_in_rm_column] == '2х2' \
                    #     else 4 if r[people_in_rm_column] == '3х3' \
                    #     else 1 if timesmena_column != 0 \
                    #     else 1
                    woman_in_rm = r[woman_in_rm_column] if woman_in_rm_column != 0 \
                        else 0

                    rm_dict = {'caption': wp_name,
                               'analog': 0,
                               'oborud': oborud,
                               'material': material,
                               'fio': [fio, ],
                               'snils': [snils, ],
                               'ind_code': [ind_code, ],
                               'woman_in_rm': [woman_in_rm, ],
                               'rm_type': rm_type,
                               'etks': etks,
                               'codeok': codeok,
                               'address': address,
                               'timesmena': timesmena,
                               'people_in_rm': people_in_rm}

                    # Создание ключа рабочих мест
                    if 'rm' not in current:
                        current['rm'] = [rm_dict, ]

                    # Если такого рабочего места нет — создаём новую запись
                    elif r[count_column].isdigit():
                        current['rm'].append(rm_dict)

                    # Если место есть — добавляем аналогию
                    else:
                        current['rm'][-1]['analog'] += 1
                        current['rm'][-1]['fio'].append(fio)
                        current['rm'][-1]['snils'].append(snils)
                        current['rm'][-1]['ind_code'].append(ind_code)
                        current['rm'][-1]['woman_in_rm'].append(woman_in_rm)

        print(f'Создан json файл на {wp_count} рм')
        return wp_count
