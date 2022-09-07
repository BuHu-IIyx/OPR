from import_in_access.att_main import make_json_org, import_in_db, create_template_for_template
from prof_risks.risk_main import import_factors_from_csv, import_organization_csv, create_cards_docx

if __name__ == '__main__':
    make_json_org('ПАО ВТБ 2 этап')
    import_in_db('ПАО ВТБ 2 этап')
    # create_template_for_template()
    # import_factors_from_csv('C:\\Users\\buhu_\\PycharmProjects\\OPR\\input\\for_risks\\risks_dic.csv')
    # import_organization_csv('C:\\Users\\buhu_\\PycharmProjects\\OPR\\input\\for_risks\\rms.csv', 'Василек',
    #                         'Генеральный директор', 'Васильев В.В.')
    # create_cards_docx()
