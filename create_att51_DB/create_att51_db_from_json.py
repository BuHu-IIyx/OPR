import datetime
import os
import random

import pyodbc
import json
import shutil

from create_att51_DB.DB_connection import DBConnector


class CreateDB:
    def __init__(self, org_name, template_name):
        self.org_name = org_name
        self.template_name = template_name
        conn_str = f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\{org_name}'
        if not os.path.exists(conn_str):
            os.mkdir(conn_str)
        conn_str += f'\\{org_name}_БД'
        if os.path.exists(conn_str):
            shutil.rmtree(conn_str, ignore_errors=True)
        shutil.copytree(r'C:\Users\buhu_\PycharmProjects\OPR\data\NewDB', conn_str)
        json_file_address = f'output\\{org_name}\\dict.json'

        self.conn_str = conn_str
        self.db = DBConnector(conn_str)

        # Загружаем файл с SQL командами:
        sql_dict_json_file = 'C:\\Users\\buhu_\\PycharmProjects\\OPR\\create_att51_DB\\data\\sql_dict.json'
        with open(sql_dict_json_file, 'r', encoding="utf-8-sig") as file:
            self.sql_dict = json.load(file)

        # Загружаем файл с шаблонами рабочих мест:
        dict_json_file = f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\Templates\\{self.template_name}\\dict.json'
        with open(dict_json_file, 'r', encoding="utf-8-sig") as file:
            self.rm_dict = json.load(file)

        # Загружаем файл со структурой организации:
        with open(json_file_address, 'r', encoding="utf-8-sig") as file:
            self.org_struct = json.load(file)

        # Получаем максимальное значение m_order из БД:
        self.max_org = self.db.select_one_from_DB(self.sql_dict['select']['max_order'] + 'struct_org')
        self.max_ceh = self.db.select_one_from_DB(self.sql_dict['select']['max_order'] + 'struct_ceh')
        self.max_uch = self.db.select_one_from_DB(self.sql_dict['select']['max_order'] + 'struct_uch')
        self.max_rm = self.db.select_one_from_DB(self.sql_dict['select']['max_order'] + 'struct_rm')
        self.wp_count = 0

    @staticmethod
    def get_mguid(name_id: str, m_order=1) -> str:
        unique_num = int((datetime.datetime.now().timestamp() % 10) * 1000000000000000)
        return hex(unique_num * m_order + id(name_id))[2:].upper()

    def insert_ceh(self, id_org):
        # Перебираем цеха:
        for ceh in self.org_struct['ceh'].keys():
            # Создаём кортеж для вставки и вставляем его в БД:
            ceh_tuple = (id_org, ceh, 0, self.max_ceh, self.get_mguid(ceh, self.max_ceh))
            self.db.insert_in_DB(self.sql_dict['insert']['add_ceh'], ceh_tuple)

            # Получение id нового цеха:
            id_ceh = self.db.select_one_from_DB(self.sql_dict['select']['id_last'])
            self.max_ceh += 1

            # Вызов функции вставки участков:
            self.insert_uch(self.org_struct['ceh'][ceh]['uch'], id_ceh)

    def insert_uch(self, uchs, id_ceh, id_par=0):
        # Перебираем участки в цехе(участке):
        for uch in uchs.keys():
            # Если встречаем рабочее место вызываем вставку РМ в БД:
            if uch == 'rm':
                self.insert_rm(uchs['rm'], id_ceh, id_par)
            # Если это участок, вставляем и вызываем вставку дочерних участков:
            else:
                node_level = 0 if id_par == 0 else 1

                # Создаём кортеж для вставки и вставляем его в БД:
                uch_tuple = (id_par, id_ceh, node_level, uch, 0, self.max_uch, self.get_mguid(uch, self.max_ceh))
                self.db.insert_in_DB(self.sql_dict['insert']['add_uch'], uch_tuple)

                # Получение id нового участка:
                id_uch = self.db.select_one_from_DB(self.sql_dict['select']['id_last'])
                self.max_uch += 1

                # Вызов функции вставки дочерних участков:
                self.insert_uch(uchs[uch], id_ceh, id_uch)

    def insert_rm(self, rms, id_ceh, id_uch):
        for rm in rms:
            # Создаём уникальный идентификатор:
            mguid = self.get_mguid(rm['caption'], self.max_ceh)
            # Создаём кортеж для вставки и вставляем его в БД:
            kut1 = 2
            file_sout = mguid + r'\Карта СОУТ.docx'
            codeok = rm['codeok'] if rm['codeok'] != '' else self.rm_dict[rm['rm_type']]['rm_data']['codeok']
            etks = rm['etks'] if rm['codeok'] != '' else self.rm_dict[rm['rm_type']]['rm_data']['etks']
            rm_tuple = (rm['caption'], id_ceh, id_uch, codeok, etks, self.max_rm, mguid, kut1, file_sout,
                        str(rm['ind_code'][0]), rm['address'], rm['timesmena'], rm['people_in_rm'])
            self.db.insert_in_DB(self.sql_dict['insert']['add_rm'], rm_tuple)

            # Получение id нового РМ:
            id_rm = self.db.select_one_from_DB(self.sql_dict['select']['id_last'])
            # Вставка ФИО и СНИЛС:
            rabs_tuple = (id_rm, rm['fio'][0], rm['snils'][0])
            self.db.insert_in_DB(self.sql_dict['insert']['sout_rabs'], rabs_tuple)
            # Вставка данных из шаблона
            self.insert_other_rm_data(rm['rm_type'], id_rm, mguid, rm)
            self.max_rm += 1
            # ПОЧЕМУ ТЫ НЕ РАБОТАЕШЬ!!!
            # if len(rm['oborud']) > 0:
            #     self.db.insert_in_DB(self.sql_dict['update']['sout_dop_info_fact'], rm['oborud'], rm['material'], id_rm)
            # Создание файлов по шаблону:
            self.insert_data(mguid, rm['rm_type'])

            # Если у РМ есть аналогичные добавляем их:
            if rm['analog'] > 0:
                self.insert_analog(id_rm, id_ceh, id_uch, kut1, rm)

    def insert_analog(self, id_rm, id_ceh, id_uch, kut1, rm):
        analog_count = rm['analog'] + 1
        count = (analog_count * 0.2) - 0.1 if analog_count >= 10 else 2
        # Добавляем группу аналогичности:
        anal_tuple = (0, id_rm, id_rm)
        self.db.insert_in_DB(self.sql_dict['insert']['add_analog'], anal_tuple)

        # Получение id новой группы аналогичности и вставка его:
        anal_id = self.db.select_one_from_DB(self.sql_dict['select']['id_last'])
        self.db.insert_in_DB(self.sql_dict['update']['update_analog'], anal_id, anal_id)

        # Добавляем все аналогичные РМ:
        for i in range(rm['analog']):
            # Создаём новый уникальный идентификатор:
            mguid = self.get_mguid(str(i), self.max_ceh)
            rm_mguid = 0
            # Создаём кортеж для вставки и вставляем его в БД:
            file_sout = ''
            codeok = rm['codeok'] if rm['codeok'] != '' else self.rm_dict[rm['rm_type']]['rm_data']['codeok']
            etks = rm['etks'] if rm['codeok'] != '' else self.rm_dict[rm['rm_type']]['rm_data']['etks']
            rm_tuple = (rm['caption'], id_ceh, id_uch, codeok, etks, self.max_rm, mguid, kut1, file_sout,
                        str(rm['ind_code'][i + 1]), rm['address'], rm['timesmena'], rm['people_in_rm'])
            self.db.insert_in_DB(self.sql_dict['insert']['add_rm'], rm_tuple)
            # Получение id нового РМ и вставка в группу аналогии:
            id_anal_rm = self.db.select_one_from_DB(self.sql_dict['select']['id_last'])
            # Вставка ФИО и СНИЛС:
            an_rabs_tuple = (id_anal_rm, rm['fio'][i + 1], rm['snils'][i + 1])
            self.db.insert_in_DB(self.sql_dict['insert']['sout_rabs'], an_rabs_tuple)
            # Вставка данных из шаблона
            if count > 1:
                self.insert_data(mguid, rm['rm_type'])
                count -= 1
                rm_mguid = mguid

            self.insert_other_rm_data(rm['rm_type'], id_anal_rm, rm_mguid, rm)

            # if len(rm['oborud']) > 0:
            #     self.db.insert_in_DB(self.sql_dict['update']['sout_dop_info_fact'], rm['oborud'], rm['material'],
            #                          id_anal_rm)
            self.max_rm += 1
            self.db.insert_in_DB(self.sql_dict['insert']['add_analog'], anal_id, id_anal_rm, id_rm)

    def insert_other_rm_data(self, rm_type, id_rm, rm_mguid, rm):
        rm_keys_dict = ["per_genfactors", "sout_dop_info2", "sout_dop_info_fact", "sout_dop_info_norm", "sout_factors",
                        "sout_ident", "sout_karta_dop_info"]
        rzone_keys_dict = ["per_gigfactors", "per_rzona_mat"]
        # Вставляем данные о рабочих местах
        for key in rm_keys_dict:
            if key in self.rm_dict[rm_type].keys():
                if key == 'sout_dop_info_fact' and rm['oborud'] != '' and rm['material'] != '':
                    for i in self.rm_dict[rm_type][key]:
                        other_tuple = (id_rm, rm['oborud'], rm['material'], *i[2:])
                        self.db.insert_in_DB(self.sql_dict['insert'][key], other_tuple)

                elif key == 'sout_factors':
                    if rm_mguid != 0:
                        for i in self.rm_dict[rm_type][key]:
                            tmp = i[-1].split('\\')[-1]
                            address = rm_mguid + '\\' + tmp
                            mguid = self.get_mguid(str(i[0]), self.max_ceh)
                            other_tuple = (id_rm, *i[:-1], address, mguid)
                            self.db.insert_in_DB(self.sql_dict['insert'][key], other_tuple)
                            self.db.insert_in_DB(self.sql_dict['insert']['sout_factor_info'], mguid)
                else:
                    for i in self.rm_dict[rm_type][key]:
                        other_tuple = (id_rm, *i)
                        self.db.insert_in_DB(self.sql_dict['insert'][key], other_tuple)
        # Добавляем рабочие зоны и дополняем их данными
        if 'per_rzona' in self.rm_dict[rm_type].keys():
            for i in self.rm_dict[rm_type]['per_rzona']:
                other_tuple = (id_rm, *i)
                self.db.insert_in_DB(self.sql_dict['insert']['per_rzona'], other_tuple)
                # Получаем ID рабочего места
                rzone_id = self.db.select_one_from_DB(self.sql_dict['select']['id_last'])
                # Вставляем данные зон
                os_skor = random.randint(1, 2) / 10
                os_patm = random.randint(758, 760)
                os_temp = random.randint(200, 220) / 10
                os_vlag = random.randint(380, 420) / 10
                tuple_zone_info = (rzone_id, os_skor, os_patm, os_temp, os_vlag)
                self.db.insert_in_DB(self.sql_dict['insert']['per_rzona_info'], tuple_zone_info)
                for key in rzone_keys_dict:
                    if key in self.rm_dict[rm_type].keys():
                        for j in self.rm_dict[rm_type][key]:
                            other_tuple = (rzone_id, *j)
                            self.db.insert_in_DB(self.sql_dict['insert'][key], other_tuple)
        return

    # Добавление новой организации
    def create_DB(self):
        org_tuple = (self.org_name, self.max_org, self.get_mguid(str(self.org_name), self.max_org))
        self.db.insert_in_DB(self.sql_dict['insert']['add_org'], org_tuple)
        org_id = self.db.select_one_from_DB(self.sql_dict['select']['id_last'])
        self.insert_ceh(org_id)

    # Копирование файлов из шаблона
    def insert_data(self, mguid, type_rm):
        if type_rm in self.rm_dict.keys():
            dist = self.conn_str + '\\ARMv51_files\\' + mguid
            p_dir = f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\Templates\\{self.template_name}\\{type_rm}'
            all_folders = os.listdir(p_dir)
            rand_int = random.randint(0, len(all_folders) - 1)
            p_dir += '\\' + all_folders[rand_int]
            shutil.copytree(p_dir, dist)
            # Подсчет рабочих мест в оплату
            self.wp_count += 1
        else:
            print('Для ' + type_rm + 'не создан шаблон!!!')

    def __del__(self):
        print(f'{self.org_name}: Создана БД на {self.wp_count} рм в оплату.')
    #     # self.db.__del__()
