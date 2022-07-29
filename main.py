import datetime
import pyodbc
import csv
import shutil


def get_mguid(caption, m_order=1):
    unique_num = int((datetime.datetime.now().timestamp() % 10) * 1000000000000000)
    return hex(unique_num * m_order + id(caption))[2:].upper()


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
    sql_str = 'INSERT INTO struct_ceh(org_id, caption, deleted, m_order, mguid) VALUES(?, ?, 0, ?, ?)'
    mguid = get_mguid(capt)
    try:
        cursor = con.cursor()
        cursor.execute(sql_str, orgn, capt, order, mguid)

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
              '(`par_id`, `ceh_id`, `node_level`, `caption`, `deleted`, `m_order`, `mguid`) ' \
              'VALUES(?, ?, ?, ?, 0, ?, ?)'
    mguid = get_mguid(caption)
    try:
        cursor = con.cursor()
        cursor.execute(sql_str, par_id, ceh_id, node_level, caption, m_order, mguid)

    except pyodbc.Error as e:
        print("Error in Connection", e)


def get_rm_id(con, ceh_id, uch_id, caption, is_main=False):
    sql_str = 'SELECT id FROM struct_rm WHERE ceh_id = ? AND uch_id = ? AND caption = ? ORDER BY id'
    try:
        cursor = con.cursor()
        cursor.execute(sql_str, ceh_id, uch_id, caption)
        row = cursor.fetchall()
        if row:
            if is_main:
                return row[0][0]
            else:
                return row[len(row)-1][0]
        else:
            return -1

    except pyodbc.Error as e:
        print("Error in Connection", e)


def add_rm(con, ceh_id, uch_id, caption, codeok, etks, m_order, is_main=True):
    sql_str = 'INSERT INTO struct_rm (caption, ceh_id, uch_id, codeok, etks, m_order, mguid, kut1, file_sout) ' \
              'VALUES(?, ?, ?, ?, ?, ?, ?, 2, ?)'
    mguid = get_mguid(caption, m_order)
    file_sout = (mguid + '\Карта СОУТ.docx') if is_main else ''
    if is_main:
        dist = 'C:\\Users\\buhu_\\PycharmProjects\\OPR\\DB\\ARMv51_files\\'
        shutil.copytree(r'C:\Users\buhu_\PycharmProjects\OPR\office', dist + mguid)
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


def import_in_att(db_address, csv_address, org, column):
    conn = start_connect(db_address)
    max_ceh = get_max_order(conn, 'struct_ceh') + 1
    max_uch = get_max_order(conn, 'struct_uch') + 1
    max_wp = get_max_order(conn, 'struct_rm') + 1
    with open(csv_address, newline='', encoding="utf-8") as File:
        reader = csv.reader(File, delimiter=';')
        flag = True
        for r in reader:
            ceh = None
            current_uch = 0
            for i in range(column):
                if len(r[i]) == 0:
                    break
                s = r[i][0].upper() + r[i][1:]
                if i == 0:
                    if check_ceh(conn, s, org) == -1:
                        add_ceh(conn, org, s, max_ceh)
                        max_ceh += 1
                    ceh = check_ceh(conn, s, org)
                else:
                    node = 0 if i == 1 else 1
                    if check_uch(conn, s, current_uch, ceh) == -1:
                        add_uch(conn, current_uch, ceh, node, s, max_uch)
                        max_uch += 1
                    current_uch = check_uch(conn, s, current_uch, ceh)
            # wp_name = r[7][0].upper() + r[7][1:]
            # if get_rm_id(conn, ceh, current_uch, wp_name) == -1:
            #     add_rm(conn, ceh, current_uch, wp_name, r[8], r[9], max_wp)
            #     flag = True
            # else:
            #     add_rm(conn, ceh, current_uch, wp_name, r[8], r[9], max_wp, False)
            #     main_rm = get_rm_id(conn, ceh, current_uch, wp_name, True)
            #     current_rm = get_rm_id(conn, ceh, current_uch, wp_name)
            #     if flag:
            #         add_analog(conn, main_rm, main_rm)
            #     add_analog(conn, current_rm, main_rm)
            #     flag = False
            # max_wp += 1

    conn.commit()
    conn.close()


if __name__ == '__main__':
    db_conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};' \
                  r'DBQ=C:\Users\buhu_\PycharmProjects\OPR\DBksu\ARMv51.MDB;'
    # csv_address_str = 'First.csv'
    # organization = 1
    # csv_address_str = 'Second.csv'
    # organization = 2
    csv_address_str = 'DBksu/ksu.csv'
    organization = 1
    import_in_att(db_conn_str, csv_address_str, organization, 4)
