import pyodbc
import csv


def get_max_order(con, table):
    sql_str = f'SELECT MAX(m_order) FROM {table}'
    try:
        cursor = con.cursor()
        cursor.execute(sql_str)
        row = cursor.fetchone()
        if row[0] is not None:
            return row[0]
        else:
            return 0

    except pyodbc.Error as e:
        print("Error in Connection", e)


# def get_max_uch(con):
#     sql_str = 'SELECT MAX(m_order) FROM struct_uch'
#     try:
#         cursor = con.cursor()
#         cursor.execute(sql_str)
#         row = cursor.fetchone()
#         if row[0] is not None:
#             return row[0]
#         else:
#             return 0
#
#     except pyodbc.Error as e:
#         print("Error in Connection", e)


def check_ceh(con, name, org_id) -> int:
    sql_str = 'SELECT id FROM struct_ceh WHERE caption = ? AND org_id = ?'
    try:
        cursor = con.cursor()
        cursor.execute(sql_str, name, org_id)
        row = cursor.fetchone()
        if row:
            return row[0]
        else:
            return -1

    except pyodbc.Error as e:
        print("Error in Connection", e)


def add_ceh(con, orgn, capt, order):
    sql_str = 'INSERT INTO struct_ceh(org_id, caption, deleted, m_order) VALUES(?, ?, 0, ?)'
    try:
        cursor = con.cursor()
        cursor.execute(sql_str, orgn, capt, order)

    except pyodbc.Error as e:
        print("Error in Connection", e)


def check_uch(con, name, par_id, ceh_id):
    sql_str = 'SELECT id FROM struct_uch WHERE caption = ? AND par_id = ? AND ceh_id = ?'
    try:
        cursor = con.cursor()
        cursor.execute(sql_str, name, par_id, ceh_id)
        row = cursor.fetchone()
        if row:
            return row[0]
        else:
            return -1

    except pyodbc.Error as e:
        print("Error in Connection", e)


def add_uch(con, par_id, ceh_id, node_level, caption, m_order):
    sql_str = 'INSERT INTO `struct_uch` ' \
              '(`par_id`, `ceh_id`, `node_level`, `caption`, `deleted`, `m_order`) ' \
              'VALUES(?, ?, ?, ?, 0, ?)'
    try:
        cursor = con.cursor()
        cursor.execute(sql_str, par_id, ceh_id, node_level, caption, m_order)

    except pyodbc.Error as e:
        print("Error in Connection", e)


def get_rm_id(con, ceh_id, uch_id, caption):
    sql_str = 'SELECT MAX(id) FROM struct_rm WHERE ceh_id = ? AND uch_id = ? AND caption = ?'
    try:
        cursor = con.cursor()
        cursor.execute(sql_str, ceh_id, uch_id, caption)
        row = cursor.fetchone()
        if row:
            return row[0]
        else:
            return -1

    except pyodbc.Error as e:
        print("Error in Connection", e)


def add_rm(con, ceh_id, uch_id, caption, codeok, etks, m_order):
    sql_str = 'INSERT INTO struct_rm (caption, ceh_id, uch_id, codeok, etks, m_order, mguid, kut1, file_sout) ' \
              'VALUES(?, ?, ?, ?, ?, ?, ?, 2, ?)'
    mguid = 'pass'
    file_sout = mguid + '\Карта СОУТ.docx'
    try:
        cursor = con.cursor()
        cursor.execute(sql_str, caption, ceh_id, uch_id, codeok, etks, m_order, mguid, file_sout)

    except pyodbc.Error as e:
        print("Error in Connection", e)


def get_analog_id(con, rm_id):
    sql_str = 'SELECT id FROM anal_group WHERE rm_id = ? AND main = ?'
    try:
        cursor = con.cursor()
        cursor.execute(sql_str, rm_id, rm_id)
        row = cursor.fetchone()
        if row:
            return row[0]
        else:
            return -1

    except pyodbc.Error as e:
        print("Error in Connection", e)


def add_analog(con, rm_id, main):
    sql_str = 'INSERT INTO anal_group (group_id, rm_id, main) ' \
              'VALUES(?, ?, ?)'
    update_str = 'UPDATE anal_group SET group_id = ? WHERE rm_id = ? AND main = ?'
    get_str = 'SELECT id FROM anal_group WHERE rm_id = ? AND main = ?'
    if rm_id == main:
        group_id = 0
        try:
            cursor = con.cursor()
            cursor.execute(sql_str, group_id, rm_id, main)
            cursor.commit()
            group_id = cursor.execute(get_str, rm_id, main).fetchone()[0]
            cursor.execute(update_str, group_id, rm_id, main)
        except pyodbc.Error as e:
            print("Error in Connection", e)
    else:
        try:
            cursor = con.cursor()
            group_id = cursor.execute(get_str, main, main).fetchone()[0]
            cursor.execute(sql_str, group_id, rm_id, main)

        except pyodbc.Error as e:
            print("Error in Connection", e)


def start_connect(conn_str):
    try:
        connect = pyodbc.connect(conn_str)
        print("Connected To Database")
        return connect

    except pyodbc.Error as e:
        print("Error in Connection", e)


def import_in_att(db_address, csv_address):
    conn = start_connect(db_address)
    org = 1
    max_ceh = get_max_order(conn, 'struct_ceh') + 1
    max_uch = get_max_order(conn, 'struct_uch') + 1
    max_wp = get_max_order(conn, 'struct_rm') + 1
    with open(csv_address, newline='', encoding="utf-8") as File:
        reader = csv.reader(File, delimiter=';')
        rm_prev = [0, '']
        flag = True
        for r in reader:
            ceh = None
            current_uch = 0
            node = 0
            for i in range(7):
                if len(r[i]) == 0:
                    break
                s = r[i][0].upper() + r[i][1:]
                if i == 0:
                    if check_ceh(conn, s, org) == -1:
                        add_ceh(conn, org, s, max_ceh)
                        max_ceh += 1
                    ceh = check_ceh(conn, s, org)
                else:
                    if check_uch(conn, s, current_uch, ceh) == -1:
                        add_uch(conn, current_uch, ceh, node, s, max_uch)
                        max_uch += 1
                        node = 1
                    current_uch = check_uch(conn, s, current_uch, ceh)
            add_rm(conn, ceh, current_uch, r[7], r[8], r[9], max_wp)
            if r[7] == rm_prev[1]:
                if flag:
                    print(rm_prev)
                    add_analog(conn, rm_prev[0], rm_prev[0])
                add_analog(conn, get_rm_id(conn, ceh, current_uch, r[7]), rm_prev[0])
                flag = False
            else:
                rm_prev = [get_rm_id(conn, ceh, current_uch, r[7]), r[7]]
                flag = True
    conn.commit()
    conn.close()


if __name__ == '__main__':
    db_conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};' \
                  r'DBQ=C:\Users\buhu_\PycharmProjects\OPR\ARMv51.MDB;'
    csv_address_str = 'First.csv'
    import_in_att(db_conn_str, csv_address_str)
