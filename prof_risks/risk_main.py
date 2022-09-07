import csv
from sqlalchemy import create_engine, select, MetaData, Table, insert


from docx.enum.section import WD_ORIENTATION
from docx import Document
from docx.shared import Mm, Cm, Pt
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import ns
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml


def import_factors_from_csv(csv_file):
    engine = create_engine('sqlite:///prof_risks/sqlite.db')
    metadata = MetaData(engine)
    conn = engine.connect()
    dangers_type = Table('dangers_type', metadata, autoload=True)
    dangers_dict = Table('dangers_dict', metadata, autoload=True)

    with open(csv_file, newline='', encoding="utf-8-sig") as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            risk_type = conn.execute(select([dangers_type.c.id]).where(dangers_type.c.name == row[1])).fetchone()
            if risk_type is None:
                risk_type = conn.execute(insert(dangers_type).values(name=row[1])).inserted_primary_key
            conn.execute(insert(dangers_dict).values(id=row[0], description=row[2], type=risk_type[0]))


def import_organization_csv(csv_file, org_name, org_head_pos='', org_head_name=''):
    engine = create_engine('sqlite:///prof_risks/sqlite.db')
    metadata = MetaData(engine)
    conn = engine.connect()
    organizations = Table('organizations', metadata, autoload=True)
    departments = Table('departments', metadata, autoload=True)
    work_places = Table('work_places', metadata, autoload=True)
    result = Table('result', metadata, autoload=True)
    dangers_group = Table('dangers_group', metadata, autoload=True)
    insert_org_sql = insert(organizations).values(name=org_name, head_position=org_head_pos, head_name=org_head_name)
    org_id = conn.execute(insert_org_sql).inserted_primary_key[0]
    with open(csv_file, newline='', encoding="utf-8-sig") as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            dep_id = conn.execute(select([departments.c.id]).where(departments.c.name == row[1])).fetchone()
            if dep_id is None:
                insert_dep_sql = insert(departments).values(name=row[1], address='', organization_id=org_id)
                dep_id = conn.execute(insert_dep_sql).inserted_primary_key
            insert_wp = insert(work_places).values(name=row[0], equipment=row[2], department_id=dep_id[0])
            wp_id = conn.execute(insert_wp).inserted_primary_key[0]
            wp_template = conn.execute(select([dangers_group]).where(dangers_group.c.name == row[7]))
            for risk in wp_template:
                insert_result_sql = insert(result).values(work_place_id=wp_id, danger_id=risk[2], probability=risk[3],
                                                          severity=risk[4], comment=risk[5], measures=risk[6])
                conn.execute(insert_result_sql)


def create_cards_docx():
    # Create and customization document
    doc = Document()
    current_section = doc.sections[-1]
    new_width, new_height = current_section.page_height, current_section.page_width
    current_section.orientation = WD_ORIENTATION.LANDSCAPE
    current_section.page_width = new_width
    current_section.page_height = new_height
    current_section.left_margin = Cm(2)
    current_section.right_margin = Cm(2)
    current_section.top_margin = Cm(3)
    current_section.bottom_margin = Mm(15)
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(10)

    doc.add_paragraph().add_run("Раздел 4. Оценка профессиональных рисков").bold = True
    table = doc.add_table(rows=3, cols=11)
    table.style = 'Table Grid'
    table_cells = table.rows[0].cells
    table.columns[0].width = Mm(5)
    table_cells[0].width = Mm(5)
    table.cell(0, 0).merge(table.cell(2, 0)).paragraphs[0].add_run("№\nп/п").font.size = Pt(9)
    table.cell(0, 1).merge(table.cell(2, 1)).paragraphs[0].add_run("Наименование\nпрофессии").font.size = Pt(9)
    table.cell(0, 2).merge(table.cell(2, 2)).paragraphs[0].add_run("Объект оценки\nпрофессиональных\nрисков")\
        .font.size = Pt(9)
    table.cell(0, 3).merge(table.cell(2, 3)).paragraphs[0].add_run("Идентификация\nопасностей (код "
                                                                   "и\nнаименование\nопасности)").font.size = Pt(9)
    table.cell(0, 4).merge(table.cell(0, 8)).paragraphs[0].add_run("Индекс профессионального риска (ИПР)")\
        .font.size = Pt(9)
    table.cell(1, 4).merge(table.cell(1, 5)).paragraphs[0].add_run("Вероятность").font.size = Pt(9)
    table.cell(1, 6).merge(table.cell(1, 7)).paragraphs[0].add_run("Тяжесть ущерба").font.size = Pt(9)
    table.cell(2, 4).paragraphs[0].add_run("Оценка").font.size = Pt(9)
    table.cell(2, 5).paragraphs[0].add_run("Балл").font.size = Pt(9)
    table.cell(2, 6).paragraphs[0].add_run("Оценка").font.size = Pt(9)
    table.cell(2, 7).paragraphs[0].add_run("Балл").font.size = Pt(9)
    table.cell(1, 8).merge(table.cell(2, 8)).paragraphs[0].add_run("Итог").font.size = Pt(9)
    table.cell(0, 9).merge(table.cell(2, 9)).paragraphs[0].add_run("Комментарии\nаудитора").font.size = Pt(9)
    table.cell(0, 10).merge(table.cell(2, 10)).paragraphs[0].add_run("Корректирующие\nмероприятия").font.size = Pt(9)

    for i in range(0, 3):
        for j in range(0, 11):
            cell = table.cell(i, j)
            cell.paragraphs[0].alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            shading_elm = parse_xml(r'<w:shd {} w:fill="DEEAF6"/>'.format(nsdecls('w')))
            cell._tc.get_or_add_tcPr().append(shading_elm)

    file_name = f'Output/test_cards.docx'
    doc.save(file_name)


def create_report_docx():
    pass

