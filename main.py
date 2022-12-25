import os

from create_att51_DB.DB_connection import DBConnector
from create_att51_DB.create_att51_template_from_DB import CreateTemplate
from create_att51_DB.create_att51_db_from_json import CreateDB
from create_wp_json.create_wp_json_from_csv import make_json_org
from prof_risks.DBAdapter import DBAdapter
from prof_risks.DocxAdapter import DocxAdapter


# from import_in_access.att_main import make_json_org, import_in_db, create_template_for_template
# from prof_risks.risk_main import import_templates_from_csv, import_organization_json, generate_cards, \
#     fill_table_risk_evaluation, create_docx_from_template


def create_db(name, template_name):
    make_json_org(name)
    db_interface = CreateDB(name, template_name)
    db_interface.create_DB()


def update_db():
    db_path = "C:\\Users\\buhu_\\Работа\\Оборонка\\базы"
    all_folders = os.listdir(db_path)
    dict_measure = {
        'Тяжесть (ж)': ['Тяжесть: Организовать рациональные режимы труда  и отдыха',
                        'Снижение тяжести трудового процесса'],
        'Тяжесть (м)': ['Тяжесть: Организовать рациональные режимы труда  и отдыха',
                        'Снижение тяжести трудового процесса'],
        'Аэрозоли ПФД': ['Аэрозоли ПФД: Организовать рациональные режимы труда  и отдыха',
                         'Уменьшение времени контакта с вредными веществами'],
        'Химический': ['Химический: Организовать рациональные режимы труда  и отдыха',
                       'Уменьшение времени контакта с вредными веществами'],
        'Микроклимат': ['Микроклимат: Организовать рациональные режимы труда  и отдыха',
                        'Снижение времени воздействия фактора'],
        'Шум': ['Шум: Определить необходимость оснащения рабочего места средствами индивидуальной защиты органов слуха;'
                ' определить необходимость применения технических средств по снижению уровней шума в соответствии с'
                ' требованиями СанПиН 1.2.3685-21 и рассмотреть вопрос ограничения времени воздействия шума на рабочих'
                ' в соответствии с Р 2.2.2006-05', 'Снижение воздействия вредного фактора на организм человека'],
        'УФ-излучение': ['УФ-излучение: Контроль за состоянием здоровья, с целью выявления профессиональных '
                         'заболеваний',
                         'Проведение медицинского осмотра']
    }
    for folder in all_folders:
        db_conn = DBConnector(db_path + '\\' + folder)
        sql = "SELECT id, factor_name FROM sout_factors WHERE KUT='3.1' OR KUT='3.2'"
        res = db_conn.execute_DB1(sql)
        sql_update = 'UPDATE sout_factor_info SET measure=?, purpose=? WHERE id=?'
        for rm in res:
            fact_id = rm[0]
            fact_name = rm[1]
            db_conn.insert_in_DB(sql_update, dict_measure[fact_name][0], dict_measure[fact_name][1], fact_id)
        print(folder)


if __name__ == '__main__':
    # update_db()
    # conn_str = 'C:\\Users\\buhu_\\PycharmProjects\\OPR\\input\\template'

    # Создание контингента:
    # conn_str1 = 'C:\\Users\\buhu_\\Работа\\СБЕР\\SBER'
    # template = CreateTemplate(conn_str1)
    # template.create_template('SBER')

    # Создание базы данных:
    # arr_sber = ['СБЕР Вологда', 'СБЕР Головной', 'СБЕР Калининград', 'СБЕР Карелия', 'СБЕР Мурманск', 'СБЕР Псков']
    # for item in arr_sber:
    #     create_db(item, 'SBER')

    # create_att51_template()

    # Импорт контингента в риски:
    # db = DBAdapter()
    # db.import_templates_from_csv('C:\\Users\\buhu_\\PycharmProjects\\OPR\\input\\template\\Шаблоны Рисков.csv')

    # Импорт json организации в базу данных:
    org_name = 'Алроса'
    # db = DBAdapter()
    # db.import_organization_json(
    #     f'C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\РИСКИ {org_name}\\dict.json',
    #     org_name, org_head_pos='', org_head_name='')

    # Создать отчет по рискам из БД
    docAdapter = DocxAdapter(org_name, 'egida', '2022-12-369971-RDI-SC', '23.12.2022')
    docAdapter.generate_all()

    # generate_cards('Антикор', 'egida', '340501-PNT', '30.11.2022')
    # create_docx_from_template('Антикор', 'egida')

    # make_json_org('РИСКИ Монтажсервисстрой')
    # make_json_org('Риски ОСИ Внуково')
    # make_json_org('Риски ОСИ Горелово')
    # make_json_org('Риски ОСИ СПБ')

    # import_in_db('ОСИ МСК')
    # import_in_db('ОСИ Внуково')
    # import_in_db('ОСИ Горелово')
    # import_in_db('ОСИ СПБ')

    # create_template_for_template()
    # import_factors_from_csv('C:\\Users\\buhu_\\PycharmProjects\\OPR\\input\\for_risks\\risks_dic.csv')
    # import_organization_csv('C:\\Users\\buhu_\\PycharmProjects\\OPR\\input\\for_risks\\rms.csv', 'Василек',
    #                         'Генеральный директор', 'Васильев В.В.')
    # import_organization_json('C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\Риски\\ОСИ\\dict.json', 'ОСИ',
    #                          org_head_pos='', org_head_name='')
    # import_organization_json('C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\ОСИ Внуково\\dict.json', 'ОСИ внуково 2',
    #                          org_head_pos='', org_head_name='')
    # create_cards_docx('ОСИ-пробный')
    # create_docx_from_template('ОСИ', 'sercons')
    # generate_cards('ОСИ', 'sercons',  '2022-07-340613-RDI-SC')
    # fill_table_risk_evaluation('Василек')
