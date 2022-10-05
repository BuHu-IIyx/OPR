from create_att51_DB.create_att51_template_from_DB import CreateTemplate
from create_att51_DB.create_att51_db_from_json import CreateDB
from import_in_access.att_main import make_json_org, import_in_db, create_template_for_template
from prof_risks.risk_main import import_factors_from_csv, import_organization_csv, create_cards_docx, \
    fill_table_risk_evaluation, import_organization_json, create_docx_from_template, generate_cards

if __name__ == '__main__':
    # conn_str = 'C:\\Users\\buhu_\\PycharmProjects\\OPR\\input\\template'
    # template = CreateTemplate(conn_str)
    # template.create_template('template')

    DB_interface = CreateDB('ОСИ Внуково1')
    DB_interface.create_DB()
    # create_att51_template()

    # make_json_org('Риски ОСИ МСК')
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