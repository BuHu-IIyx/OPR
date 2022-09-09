import csv
import json


class ParserSCV:
    def __init__(self, csv_address: str, type='column'):
        self.csv_address = csv_address
        self.res_dict = {}
        self.type = type

    @staticmethod
    def start_with_b_l(wrong_str: str) -> str:
        return wrong_str[0].upper() + wrong_str[1:]

    def get_json(self, json_address: str):
        with open(json_address, 'w', encoding="utf-8-sig") as file:
            json.dump(self.res_dict, file, ensure_ascii=False, indent=4)

    def row_parsing(self, rm_column: int, count_column: int, fio_column=0, snils_column=0, oborud_column=0,
                    material_column=0, ind_code_column=0, rm_type_column=0, etks_column=0, codeok_column=0,
                    address_column=0):
        with open(self.csv_address, newline='', encoding="utf-8-sig") as File:
            reader = csv.reader(File, delimiter=';')
            self.res_dict['ceh'] = {}
            for r in reader:
                if len(r[0]) > 0:
                    s = self.start_with_b_l(r[1])

                    if int(r[0]) == 1:
                        current = self.res_dict['ceh']
                        if s not in current.keys():
                            current[s] = {}
                            current[s]['uch'] = {}
                        current = current[s]['uch']
                    elif int(r[0]) == 0:
                        if s not in current.keys():
                            current[s] = {}
                        current = current[s]

                elif len(r[rm_column]) > 0:
                    # Инициализация информации о рабочем месте
                    wp_name = self.start_with_b_l(r[rm_column])
                    fio = r[fio_column] if fio_column != 0 \
                        else ''
                    snils = r[snils_column] if snils_column != 0 \
                        else 'Отсутствует'
                    ind_code = r[ind_code_column] if ind_code_column != 0 \
                        else 'Отсутствует'

                    # Создание ключа рабочих мест
                    if 'rm' not in current.keys():
                        oborud = r[oborud_column] if oborud_column != 0 \
                            else 'ПЭВМ, телефон, оргтехника'
                        material = self.start_with_b_l(r[material_column]) if material_column != 0 \
                            else 'Бумага, канцелярские принадлежности'
                        rm_type = r[rm_type_column] if rm_type_column != 0 \
                            else 'office'
                        etks = r[etks_column] if etks_column != 0 \
                            else ''
                        codeok = r[codeok_column] if codeok_column != 0 \
                            else ''
                        address = f"Фактический адрес: {r[address_column]}" if address_column != 0 \
                            else ''
                        rm_dict = {'caption': wp_name,
                                   'analog': 0,
                                   'oborud': oborud,
                                   'material': material,
                                   'fio': [fio, ],
                                   'snils': [snils, ],
                                   'ind_code': [ind_code, ],
                                   'rm_type': rm_type,
                                   'etks': etks,
                                   'codeok': codeok,
                                   'address': address}
                        current['rm'] = [rm_dict, ]

                    # Если такого рабочего места нет — создаём новую запись
                    elif r[count_column].isdigit():
                        oborud = r[oborud_column] if oborud_column != 0 \
                            else 'ПЭВМ, телефон, оргтехника'
                        material = self.start_with_b_l(r[material_column]) if material_column != 0 \
                            else 'Бумага, канцелярские принадлежности'
                        rm_type = r[rm_type_column] if rm_type_column != 0 \
                            else 'office'
                        etks = r[etks_column] if etks_column != 0 \
                            else ''
                        codeok = r[codeok_column] if codeok_column != 0 \
                            else ''
                        address = f"Фактический адрес: {r[address_column]}" if address_column != 0 \
                            else ''
                        rm_dict = {'caption': wp_name,
                                   'analog': 0,
                                   'oborud': oborud,
                                   'material': material,
                                   'fio': [fio, ],
                                   'snils': [snils, ],
                                   'ind_code': [ind_code, ],
                                   'rm_type': rm_type,
                                   'etks': etks,
                                   'codeok': codeok,
                                   'address': address}
                        current['rm'].append(rm_dict)

                    # Если место есть — добавляем аналогию
                    else:
                        current['rm'][-1]['analog'] += 1
                        current['rm'][-1]['fio'].append(fio)
                        current['rm'][-1]['snils'].append(snils)
                        current['rm'][-1]['ind_code'].append(ind_code)

    def column_parsing(self, count: int, rm_column: int, count_column: int, fio_column=0, snils_column=0,
                       oborud_column=0, material_column=0, ind_code_column=0, rm_type_column=0, etks_column=0,
                       codeok_column=0, address_column=0):
        with open(self.csv_address, newline='', encoding="utf-8-sig") as File:
            reader = csv.reader(File, delimiter=';')
            self.res_dict['ceh'] = {}
            for r in reader:
                current = self.res_dict['ceh']

                for i in range(count):
                    if len(r[i]) == 0:
                        break
                    s = self.start_with_b_l(r[i])

                    if i == 0:
                        if s not in current:
                            current[s] = {}
                            current[s]['uch'] = {}
                        current = current[s]['uch']
                    else:
                        if s not in current:
                            current[s] = {}
                        current = current[s]
                # СКРЫТЬ В ВТБ
                # if address_column != 0:
                #     # Добавить адрес в подразделение:
                #     address = f"Фактический адрес: {r[address_column]}"
                #     if address not in current:
                #         current[address] = {}
                #     current = current[address]

                if rm_column != 0:
                    # Инициализация информации о рабочем месте
                    wp_name = self.start_with_b_l(r[rm_column])
                    oborud = self.start_with_b_l(r[oborud_column]) if oborud_column != 0 \
                        else 'ПЭВМ, телефон, оргтехника'
                    material = self.start_with_b_l(r[material_column]) if material_column != 0 \
                        else 'Бумага, канцелярские принадлежности'
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

                    # Создание ключа рабочих мест
                    if 'rm' not in current:
                        rm_dict = {'caption': wp_name,
                                   'analog': 0,
                                   'oborud': oborud,
                                   'material': material,
                                   'fio': [fio, ],
                                   'snils': [snils, ],
                                   'ind_code': [ind_code, ],
                                   'rm_type': rm_type,
                                   'etks': etks,
                                   'codeok': codeok,
                                   'address': address}
                        current['rm'] = [rm_dict, ]

                    # Если такого рабочего места нет — создаём новую запись
                    elif r[count_column].isdigit():
                        rm_dict = {'caption': wp_name,
                                   'analog': 0,
                                   'oborud': oborud,
                                   'material': material,
                                   'fio': [fio, ],
                                   'snils': [snils, ],
                                   'ind_code': [ind_code, ],
                                   'rm_type': rm_type,
                                   'etks': etks,
                                   'codeok': codeok,
                                   'address': address}
                        current['rm'].append(rm_dict)

                    # Если место есть — добавляем аналогию
                    else:
                        current['rm'][-1]['analog'] += 1
                        current['rm'][-1]['fio'].append(fio)
                        current['rm'][-1]['snils'].append(snils)
                        current['rm'][-1]['ind_code'].append(ind_code)
