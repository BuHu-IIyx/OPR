from import_in_access.att_main import make_json_org, import_in_db, create_template_for_template
from prof_risks.risk_main import import_factors_from_csv, import_organization_csv, create_cards_docx, \
    fill_table_risk_evaluation, import_organization_json, create_docx_from_template, generate_cards

if __name__ == '__main__':

    # make_json_org('ОСИ МСК')
    # import_in_db('ОСИ МСК')
    # make_json_org('ОСИ Внуково')
    # import_in_db('ОСИ Внуково')
    # make_json_org('ОСИ Горелово')
    # import_in_db('ОСИ Горелово')
    # make_json_org('ОСИ СПБ')
    # import_in_db('ОСИ СПБ')

    # create_template_for_template()
    # import_factors_from_csv('C:\\Users\\buhu_\\PycharmProjects\\OPR\\input\\for_risks\\risks_dic.csv')
    # import_organization_csv('C:\\Users\\buhu_\\PycharmProjects\\OPR\\input\\for_risks\\rms.csv', 'Василек',
    #                         'Генеральный директор', 'Васильев В.В.')
    # import_organization_json('C:\\Users\\buhu_\\PycharmProjects\\OPR\\output\\ОСИ Внуково\\dict.json', 'ОСИ-пробный',
    #                          org_head_pos='', org_head_name='')
    # create_cards_docx('ОСИ-пробный')
    # create_docx_from_template('ОСИ-пробный')
    generate_cards('ОСИ-пробный', 'sercons',  '2022-06-336990-ORAS-SC')
    # fill_table_risk_evaluation('Василек')