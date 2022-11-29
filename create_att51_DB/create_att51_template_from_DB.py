import os
import shutil
import json
from datetime import datetime
from create_att51_DB.DB_connection import DBConnector


class CreateTemplate:
    def __init__(self, db_path):
        sql_dict_json_file = 'C:\\Users\\buhu_\\PycharmProjects\\OPR\\create_att51_DB\\data\\create_template_new.json'
        self.db = DBConnector(db_path)
        self.res_dict = {}
        with open(sql_dict_json_file, 'r', encoding="utf-8-sig") as file:
            self.sql_dict = json.load(file)

    def create_template(self, template_name):
        rms = self.db.get_rm_from_db(template_name)
        for rm in rms:
            self.copy_DB_files(rm[1], rm[2], template_name)
            if rm[1] not in self.res_dict.keys():
                self.res_dict[rm[1]] = {}
                self.res_dict[rm[1]]['rm_data'] = {'codeok': rm[3], 'etks': rm[4]}
                for key in self.sql_dict.keys():
                    self.read_table(key, rm[0], rm[1])
        self.save_res_dict(template_name)

    def save_res_dict(self, template_name):
        if not os.path.isdir(f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\Templates\\{template_name}'):
            os.mkdir(f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\Templates\\{template_name}')
        json_address = f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\Templates\\{template_name}\\dict.json'
        try:
            with open(json_address, 'w', encoding="utf-8-sig") as file:
                json.dump(self.res_dict, file, ensure_ascii=False, indent=4)
        except TypeError as e:
            print(e)

    def read_table(self, table, rm_id, rm_name):
        rows = self.db.execute_DB(self.sql_dict[table], rm_id)
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

    def copy_DB_files(self, rm_name, path_mguid, res_folder):
        from_path = self.db.db_path + '\\ARMv51_files\\' + path_mguid
        p_dir = f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\Templates\\{res_folder}\\{rm_name}'
        dist_path = self.check_folder(p_dir)
        try:
            shutil.copytree(from_path, dist_path)
        except OSError as err:
            print('Для ' + rm_name + ' не создан шаблон!!!')
            print(err)

    def check_folder(self, path, count=0):
        res_path = path + '\\' + str(count)
        if os.path.isdir(res_path):
            return self.check_folder(path, (count + 1))
        else:
            return res_path

