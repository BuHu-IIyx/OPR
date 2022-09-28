from create_att51_template.DB_connection import DBConnector


def create_att51_template():
    conn_str = f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\input\\template'
    sql_dict_json_file = 'C:\\Users\\buhu_\\PycharmProjects\\OPR\\create_att51_template\\data\\create_template_new.json'
    output_dir = f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\Templates\\res_template.json'

    db_conn = DBConnector(conn_str, sql_dict_json_file)
    db_conn.create_template_json(output_dir)
