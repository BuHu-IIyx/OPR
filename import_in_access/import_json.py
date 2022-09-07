import datetime
import random

import pyodbc
import json
import shutil


class InterfaceDB:
    def __init__(self, conn_str, json_file_address):
        # Создаём соединение с БД:
        self.conn_str = conn_str
        db_conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}}; DBQ={conn_str}\\ARMv51.MDB;'
        try:
            self.connect = pyodbc.connect(db_conn_str)
            self.cursor = self.connect.cursor()
        except pyodbc.Error as e:
            print("Error in Connection", e)

        # Загружаем файл с SQL командами:
        sql_dict_json_file = 'C:\\Users\\buhu_\\PycharmProjects\\OPR\\data\\sql_dict.json'
        with open(sql_dict_json_file, 'r', encoding="utf-8-sig") as file:
            self.sql_dict = json.load(file)

        # Загружаем файл с шаблонами рабочих мест:
        dict_json_file = 'C:\\Users\\buhu_\\PycharmProjects\\OPR\\data\\dict.json'
        with open(dict_json_file, 'r', encoding="utf-8-sig") as file:
            self.rm_dict = json.load(file)

        # Загружаем файл со структурой организации:
        with open(json_file_address, 'r', encoding="utf-8-sig") as file:
            self.org_struct = json.load(file)

        # Получаем максимальное значение m_order из БД:
        self.max_org = self.select_one_from_DB(self.sql_dict['select']['max_order'] + 'struct_org')
        self.max_ceh = self.select_one_from_DB(self.sql_dict['select']['max_order'] + 'struct_ceh')
        self.max_uch = self.select_one_from_DB(self.sql_dict['select']['max_order'] + 'struct_uch')
        self.max_rm = self.select_one_from_DB(self.sql_dict['select']['max_order'] + 'struct_rm')

    @staticmethod
    def get_mguid(name_id: str, m_order=1) -> str:
        unique_num = int((datetime.datetime.now().timestamp() % 10) * 1000000000000000)
        return hex(unique_num * m_order + id(name_id))[2:].upper()

    def select_one_from_DB(self, sql_str, *args):
        try:
            row = self.cursor.execute(sql_str, *args).fetchone()
            if row[0] is not None:
                return row[0]
            else:
                return 0

        except pyodbc.Error as e:
            print("Error in Connection", e)

    def insert_in_DB(self, sql_str, *args):
        try:
            self.cursor.execute(sql_str, *args)

        except pyodbc.Error as e:
            print("Error in Connection", e)

    def insert_ceh(self, id_org):
        # Перебираем цеха:
        for ceh in self.org_struct['ceh'].keys():
            # Создаём кортеж для вставки и вставляем его в БД:
            ceh_tuple = (id_org, ceh, 0, self.max_ceh, self.get_mguid(ceh, self.max_ceh))
            self.insert_in_DB(self.sql_dict['insert']['add_ceh'], ceh_tuple)

            # Получение id нового цеха:
            id_ceh = self.select_one_from_DB(self.sql_dict['select']['id_last'])
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
                self.insert_in_DB(self.sql_dict['insert']['add_uch'], uch_tuple)

                # Получение id нового участка:
                id_uch = self.select_one_from_DB(self.sql_dict['select']['id_last'])
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
            rm_tuple = (rm['caption'], id_ceh, id_uch, rm['codeok'], rm['etks'], self.max_rm, mguid, kut1, file_sout,
                        str(rm['ind_code'][0]), rm['address'])
            self.insert_in_DB(self.sql_dict['insert']['add_rm'], rm_tuple)

            # Получение id нового РМ:
            id_rm = self.select_one_from_DB(self.sql_dict['select']['id_last'])
            # Вставка ФИО и СНИЛС:
            rabs_tuple = (id_rm, rm['fio'][0], rm['snils'][0])
            self.insert_in_DB(self.sql_dict['insert']['sout_rabs'], rabs_tuple)
            # Вставка данных из шаблона
            self.insert_other_rm_data(rm['rm_type'], id_rm, mguid)
            self.max_rm += 1
            self.insert_in_DB(self.sql_dict['update']['sout_dop_info_fact'], rm['oborud'], rm['material'], id_rm)
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
        self.insert_in_DB(self.sql_dict['insert']['add_analog'], anal_tuple)

        # Получение id новой группы аналогичности и вставка его:
        anal_id = self.select_one_from_DB(self.sql_dict['select']['id_last'])
        self.insert_in_DB(self.sql_dict['update']['update_analog'], anal_id, anal_id)

        # Добавляем все аналогичные РМ:
        for i in range(rm['analog']):
            # Создаём новый уникальный идентификатор:
            mguid = self.get_mguid(str(i), self.max_ceh)
            rm_mguid = 0
            # Создаём кортеж для вставки и вставляем его в БД:
            file_sout = ''
            rm_tuple = (rm['caption'], id_ceh, id_uch, rm['codeok'], rm['etks'], self.max_rm, mguid, kut1, file_sout,
                        str(rm['ind_code'][i + 1]), rm['address'])
            self.insert_in_DB(self.sql_dict['insert']['add_rm'], rm_tuple)
            # Получение id нового РМ и вставка в группу аналогии:
            id_anal_rm = self.select_one_from_DB(self.sql_dict['select']['id_last'])
            # Вставка ФИО и СНИЛС:
            an_rabs_tuple = (id_anal_rm, rm['fio'][i + 1], rm['snils'][i + 1])
            self.insert_in_DB(self.sql_dict['insert']['sout_rabs'], an_rabs_tuple)
            # Вставка данных из шаблона
            if count > 1:
                self.insert_data(mguid, rm['rm_type'])
                count -= 1
                rm_mguid = mguid

            self.insert_other_rm_data(rm['rm_type'], id_anal_rm, rm_mguid)
            self.insert_in_DB(self.sql_dict['update']['sout_dop_info_fact'], rm['oborud'], rm['material'],
                              id_anal_rm)

            self.max_rm += 1
            self.insert_in_DB(self.sql_dict['insert']['add_analog'], anal_id, id_anal_rm, id_rm)

    def insert_other_rm_data(self, rm_type, id_rm, rm_mguid):
        for key in self.rm_dict[rm_type].keys():
            if key == 'sout_factors':
                if rm_mguid != 0:
                    for i in self.rm_dict[rm_type][key]:
                        address = rm_mguid + i[-1]
                        mguid = self.get_mguid(str(i[0]), self.max_ceh)
                        other_tuple = (id_rm, *i[:-1], address, mguid)
                        self.insert_in_DB(self.sql_dict['insert'][key], other_tuple)
                        self.insert_in_DB(self.sql_dict['insert']['sout_factor_info'], mguid)

            else:
                other_tuple = (id_rm, *self.rm_dict[rm_type][key])
                self.insert_in_DB(self.sql_dict['insert'][key], other_tuple)
                # if key == 'per_rzona':
                #     id_rzone = self.select_one_from_DB(self.sql_dict['select']['id_last'])
                #     os_skor = random.randint(0, 1) / 10
                #     os_patm = random.randint(758, 760)
                #     os_temp = random.randint(200, 220) / 10
                #     os_vlag = random.randint(380, 420) / 10
                #     self.insert_in_DB(self.sql_dict['insert']['per_rzona_info'], id_rzone, os_skor, os_patm, os_temp,
                #                       os_vlag)

    def insert_org(self, org_name):
        org_tuple = (org_name, self.max_org, self.get_mguid(str(org_name), self.max_org))
        self.insert_in_DB(self.sql_dict['insert']['add_org'], org_tuple)
        org_id = self.select_one_from_DB(self.sql_dict['select']['id_last'])
        return org_id

    def insert_data(self, mguid, type_rm):
        if type_rm == 'office' or type_rm == 'office_hard' or type_rm == 'office_hard_hard':
            dist = self.conn_str + '\\ARMv51_files\\' + mguid
            p_dir = f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\data\\templates\\{type_rm}'
            shutil.copytree(p_dir, dist)
        else:
            pass

    def __del__(self):
        self.connect.commit()
        self.connect.close()
