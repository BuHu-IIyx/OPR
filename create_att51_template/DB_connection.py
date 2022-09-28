import json
from datetime import datetime
import pyodbc


class DBConnector:
    def __init__(self, conn_str, sql_dict_json_file):
        self.res_dict = {}
        db_conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}}; DBQ={conn_str}\\ARMv51.MDB;'
        try:
            connect = pyodbc.connect(db_conn_str)
            self.cursor = connect.cursor()
        except pyodbc.Error as e:
            print("Error in Connection", e)

        with open(sql_dict_json_file, 'r', encoding="utf-8-sig") as file:
            self.sql_dict = json.load(file)

    def create_template_json(self, json_address):
        rms = self.get_rm_from_db('template')
        for rm in rms:
            self.res_dict[rm[1]] = {}
            for key in self.sql_dict.keys():
                self.read_table(key, rm[0], rm[1])
        self.save_res_dict(json_address)

    def save_res_dict(self, json_address):
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
        sql = 'SELECT rm.id, rm.caption FROM (struct_rm rm INNER JOIN struct_ceh ceh ON rm.ceh_id = ceh.id) ' \
              'INNER JOIN struct_org org ON ceh.org_id = org.id WHERE org.caption = ?'
        res = self.cursor.execute(sql, org_name).fetchall()
        return res

    def execute_DB(self, sql_str, args):
        try:
            res = self.cursor.execute(sql_str, args).fetchall()
            return res

        except pyodbc.Error as e:
            print("Error in Connection", e)
