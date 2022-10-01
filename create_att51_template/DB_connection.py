import json
import os
import shutil
from datetime import datetime
import pyodbc


class DBConnector:
    def __init__(self, conn_str, sql_dict_json_file):
        self.res_dict = {}
        self.db_path = conn_str
        db_conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}}; DBQ={conn_str}\\ARMv51.MDB;'
        try:
            connect = pyodbc.connect(db_conn_str)
            self.cursor = connect.cursor()
        except pyodbc.Error as e:
            print("Error in Connection", e)

        with open(sql_dict_json_file, 'r', encoding="utf-8-sig") as file:
            self.sql_dict = json.load(file)

    def create_template_json(self, template_name):
        rms = self.get_rm_from_db(template_name)
        for rm in rms:
            self.copy_DB_files(rm[1], rm[2], template_name)
            if rm[1] not in self.res_dict.keys():
                self.res_dict[rm[1]] = {}
                for key in self.sql_dict.keys():
                    self.read_table(key, rm[0], rm[1])
        self.save_res_dict(template_name)

    def save_res_dict(self, template_name):
        json_address = f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\Templates\\{template_name}\\template.json'
        try:
            with open(json_address, 'w', encoding="utf-8-sig") as file:
                json.dump(self.res_dict, file, ensure_ascii=False, indent=4)
        except TypeError as e:
            print(e)

    def read_table(self, table, rm_id, rm_name):
        rows = self.execute_DB(self.sql_dict[table], rm_id)
        # rows = self.cursor.execute(self.sql_dict[table]).fetchall()
        if len(rows) > 0:
            self.res_dict[rm_name][table] = []
            for row in rows:
                res_list = []
                for elem in row:
                    if type(elem) == datetime:
                        res_list.append(str(elem))
                    else:
                        res_list.append(elem)
                self.res_dict[rm_name][table].append(res_list)

    def get_rm_from_db(self, org_name):
        sql = 'SELECT rm.id, rm.caption, rm.mguid FROM (struct_rm rm INNER JOIN struct_ceh ceh ON rm.ceh_id = ceh.id) ' \
              'INNER JOIN struct_org org ON ceh.org_id = org.id WHERE org.caption = ?'
        res = self.cursor.execute(sql, org_name).fetchall()
        return res

    def execute_DB(self, sql_str, args):
        try:
            res = self.cursor.execute(sql_str, args).fetchall()
            return res

        except pyodbc.Error as e:
            print("Error in Connection", e)

    def copy_DB_files(self, rm_name, path_mguid, res_folder):
        dist = self.db_path + '\\ARMv51_files\\' + path_mguid
        p_dir = f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\Templates\\{res_folder}\\{rm_name}'
        res_path = self.check_folder(p_dir)
        try:
            shutil.copytree(res_path, dist)
        except OSError as err:
            print('Для ' + rm_name + 'не создан шаблон!!!')
            print(err)

    def check_folder(self, path, count=0):
        if os.path.isdir(path + str(count)):
            self.check_folder(path, (count + 1))
        else:
            return path + str(count)
