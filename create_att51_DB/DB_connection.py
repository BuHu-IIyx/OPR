import pyodbc


class DBConnector:
    def __init__(self, conn_str):
        self.res_dict = {}
        self.db_path = conn_str
        db_conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}}; DBQ={conn_str}\\ARMv51.MDB;'
        try:
            self.connect = pyodbc.connect(db_conn_str)
            self.cursor = self.connect.cursor()
        except pyodbc.Error as e:
            print("Error in Connection-1", e)

    def get_rm_from_db(self, org_name):
        sql = 'SELECT rm.id, rm.caption, rm.mguid, rm.codeok, rm.etks, rm.file_sout, rm.kut1 ' \
              'FROM (struct_rm rm INNER JOIN struct_ceh ceh ' \
              'ON rm.ceh_id = ceh.id) ' \
              'INNER JOIN struct_org org ON ceh.org_id = org.id WHERE org.caption = ? AND rm.deleted = 0'
        res = self.cursor.execute(sql, org_name).fetchall()
        return res

    def execute_DB(self, sql_str, args):
        try:
            res = self.cursor.execute(sql_str, args).fetchall()
            return res

        except pyodbc.Error as e:
            print("Error in Connection0", e)

    def execute_DB1(self, sql_str):
        try:
            res = self.cursor.execute(sql_str).fetchall()
            return res

        except pyodbc.Error as e:
            print("Error in Connection1", e, )

    def select_one_from_DB(self, sql_str, *args):
        try:
            row = self.cursor.execute(sql_str, *args).fetchone()
            if row[0] is not None:
                return row[0]
            else:
                return 0
        except pyodbc.Error as e:
            print("Error in Connection2", e)

    def insert_in_DB(self, sql_str, *args):
        try:
            self.cursor.execute(sql_str, *args)
            self.connect.commit()
            return

        except pyodbc.Error as e:
            print(sql_str, args)
            print("Error in Connection3", e)

    # def __del__(self):
    #     self.connect.commit()
    #     self.connect.close()
