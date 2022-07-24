import pyodbc
import csv


def get_max_ceh(con):
    sql_str = 'SELECT MAX(m_order) FROM struct_ceh'
    try:
        cursor = con.cursor()
        cursor.execute(sql_str)
        row = cursor.fetchone()
        if row:
            return row[0]
        else:
            return -1

    except pyodbc.Error as e:
        print("Error in Connection", e)


def get_max_uch(con):
    sql_str = 'SELECT MAX(m_order) FROM struct_uch'
    try:
        cursor = con.cursor()
        cursor.execute(sql_str)
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


def add_uch(con, par_id, ceh_id, node_level, caption, m_order):
    sql_str = 'INSERT INTO `struct_uch` ' \
              '(`par_id`, `ceh_id`, `node_level`, `caption`, `deleted`, `m_order`) ' \
              'VALUES(?, ?, ?, ?, 0, ?)'
    try:
        cursor = con.cursor()
        cursor.execute(sql_str, par_id, ceh_id, node_level, caption, m_order)

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


def start_connect():
    try:
        conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\buhu_\PycharmProjects\OPR\arm.mdb;'
        connect = pyodbc.connect(conn_str)
        print("Connected To Database")
        return connect

    except pyodbc.Error as e:
        print("Error in Connection", e)


if __name__ == '__main__':
    conn = start_connect()
    org = 28
    max_ceh = get_max_ceh(conn) + 1
    max_uch = get_max_uch(conn) + 1
    with open('Second.csv', newline='', encoding="utf-8") as File:
        reader = csv.reader(File, delimiter=';')
        for r in reader:
            ceh = None
            curent_uch = 0
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
                    if check_uch(conn, s, curent_uch, ceh) == -1:
                        add_uch(conn, curent_uch, ceh, node, s, max_uch)
                        max_uch += 1
                        node += 1
                    curent_uch = check_uch(conn, s, curent_uch, ceh)
    conn.commit()

