from sqlalchemy import create_engine, select, MetaData, Table, insert, text, and_
import json
import csv


class DBAdapter:
    def __init__(self):
        engine = create_engine('sqlite:///prof_risks/sqlite.db')
        metadata = MetaData(engine)
        self.conn = engine.connect()
        self.dangers_group = Table('dangers_group', metadata, autoload=True)
        self.dangers_type = Table('dangers_type', metadata, autoload=True)
        self.dangers_dict = Table('dangers_dict', metadata, autoload=True)
        self.organizations = Table('organizations', metadata, autoload=True)
        self.departments = Table('departments', metadata, autoload=True)
        self.work_places = Table('work_places', metadata, autoload=True)
        self.result = Table('result', metadata, autoload=True)
        self.persons = Table('persons', metadata, autoload=True)

    # ---------------------------------------------------------
    # Блок отвечающий за импорт организации в БД из JSON файла:
    # ---------------------------------------------------------
    def import_organization_json(self, json_file, org_name, org_head_pos='', org_head_name=''):
        org_id = self.insert_organization(org_name, org_head_pos, org_head_name)
        with open(json_file, 'r', encoding="utf-8-sig") as file:
            rm_dict = json.load(file)
            for ceh in rm_dict['ceh'].keys():
                self.insert_structure_of_organization(ceh, rm_dict['ceh'][ceh]['uch'], org_id)

    def insert_organization(self, org_name, org_head_pos, org_head_name):
        insert_org_sql = insert(self.organizations).values(name=org_name, head_position=org_head_pos,
                                                           head_name=org_head_name)
        return self.conn.execute(insert_org_sql).inserted_primary_key[0]

    def insert_structure_of_organization(self, curr_uch, uchs, org_id):
        for uch in uchs.keys():
            if uch == 'rm':
                dep_id = self.insert_department(curr_uch, org_id)
                for rm in uchs['rm']:
                    self.insert_work_place(rm, dep_id[0])
            else:
                temp = curr_uch + ' - ' + uch
                self.insert_structure_of_organization(temp, uchs[uch], org_id)

    def insert_department(self, department_name, org_id):
        dep_id = self.conn.execute(select([self.departments.c.id]).where(and_(
                            self.departments.c.name == department_name,
                            self.departments.c.organization_id == org_id
                        ))).fetchone()
        if dep_id is None:
            insert_dep_sql = insert(self.departments).values(name=department_name, address='', organization_id=org_id)
            dep_id = self.conn.execute(insert_dep_sql).inserted_primary_key
        return dep_id

    def insert_work_place(self, wp, dep_id):
        insert_wp = insert(self.work_places).values(name=wp['caption'], equipment=wp['oborud'], department_id=dep_id,
                                                    address=wp['address'])
        wp_id = self.conn.execute(insert_wp).inserted_primary_key[0]
        for name in wp['fio']:
            insert_fio = insert(self.persons).values(work_place_id=wp_id, name=name)
            self.conn.execute(insert_fio)
        wp_template = self.conn.execute(select([self.dangers_group]).where(self.dangers_group.c.name == wp['rm_type']))
        flag = False
        for risk in wp_template:
            flag = True
            insert_result_sql = insert(self.result).values(work_place_id=wp_id, danger_id=risk[2], probability=risk[3],
                                                           severity=risk[4], comment=risk[5], measures=risk[6],
                                                           object=risk[7], probability_after=risk[9],
                                                           severity_after=risk[10])
            self.conn.execute(insert_result_sql)
        if not flag:
            # ошибка если нет шаблона
            raise Exception(f'Шаблон для РМ {wp["caption"]} не найден!!!!!')

    # ---------------------------------------------------------
    #            Импорт CSV файла шаблонов в БД:
    # ---------------------------------------------------------
    def import_templates_from_csv(self, csv_file):
        with open(csv_file, newline='', encoding="utf-8-sig") as File:
            reader = csv.reader(File, delimiter=';')
            for row in reader:
                s = self.conn.execute(
                    select(self.dangers_group.c.id).where(
                        and_(
                            self.dangers_group.c.name == row[1],
                            self.dangers_group.c.danger_id == row[4].split()[0].strip()[:-1]
                        ))
                ).fetchone()
                if s:
                    print(row[1])
                    continue
                else:
                    self.conn.execute(insert(self.dangers_group).values(name=row[1], description=row[2],
                                                                        danger_id=row[4].split()[0].strip()[:-1],
                                                                        probability=row[6], severity=row[8],
                                                                        comment=row[13],
                                                                        measures=row[15], object=row[3],
                                                                        probability_after=row[10],
                                                                        severity_after=row[11]))

    # ---------------------------------------------------------
    #         Получение данных для заполнения карт из БД:
    # ---------------------------------------------------------
    def get_departments_from_DB(self, org_name):
        sql = text("""SELECT d.id, d.name 
            FROM departments d JOIN organizations o ON d.organization_id = o.id 
            WHERE o.name = :name""")
        dic = {'name': org_name}
        # departments = self.conn.execute(sql, **dic)
        return self.conn.execute(sql, **dic)

    def get_work_places_from_DB(self, department_id):
        sql = text("""SELECT wp.id, wp.name, wp.equipment FROM work_places wp WHERE wp.department_id = :dep_id""")
        dic = {'dep_id': department_id}
        # work_places = self.conn.execute(sql, **dic)
        return self.conn.execute(sql, **dic)

    def get_dangers_from_DB(self, wp_id):
        sql = text("""SELECT d.description, r.probability, r.severity, r.comment, r.measures, r.object 
            FROM result r JOIN dangers_dict d ON r.danger_id = d.id
            WHERE r.work_place_id = :wp_id""")
        dic = {'wp_id': wp_id}
        # dangers = self.conn.execute(sql, **dic)
        return self.conn.execute(sql, **dic)

    def get_persons_from_DB(self, wp_id):
        sql = text("""SELECT p.name FROM persons p WHERE p.work_place_id = :wp_id""")
        dic = {'wp_id': wp_id}
        # persons = self.conn.execute(sql, **dic)
        return self.conn.execute(sql, **dic)

    def get_all_wp_in_org(self, org_name):
        sql_count = text("""
                SELECT COUNT(wp.name)
                FROM work_places wp
                JOIN departments d ON wp.department_id = d.id JOIN organizations o ON d.organization_id = o.id
                WHERE o.name = :name
                """)
        dic = {'name': org_name}
        return int(self.conn.execute(sql_count, **dic).fetchone()[0])

    # ---------------------------------------------------------
    #    Получение данных для заполнения файла отчета из БД:
    # ---------------------------------------------------------
    def get_table_1_data(self, org_name):
        sql = text("""
        SELECT d.name, wp.name, dangers_dict.description, r.probability, r.severity, r.comment, r.measures, r.object
        FROM (result r JOIN dangers_dict ON r.danger_id = dangers_dict.id) 
        JOIN work_places wp ON wp.id = r.work_place_id
        JOIN departments d ON wp.department_id = d.id JOIN organizations o ON d.organization_id = o.id
        WHERE o.name = :name
        ORDER BY d.id, wp.id, r.danger_id
        """)
        sql_count = text("""
                SELECT COUNT(wp.name)
                FROM (result r JOIN dangers_dict ON r.danger_id = dangers_dict.id) 
                JOIN work_places wp ON wp.id = r.work_place_id
                JOIN departments d ON wp.department_id = d.id JOIN organizations o ON d.organization_id = o.id
                WHERE o.name = :name
                """)
        dic = {'name': org_name}
        return self.conn.execute(sql, **dic), int(self.conn.execute(sql_count, **dic).fetchone()[0])

    def get_table_2_data(self, org_name):
        sql = text("""
        SELECT d.name, wp.name, dangers_dict.description, r.probability, r.severity, r.comment, r.measures, r.object, 
        r.probability_after, r.severity_after
        FROM (result r JOIN dangers_dict ON r.danger_id = dangers_dict.id) 
        JOIN work_places wp ON wp.id = r.work_place_id
        JOIN departments d ON wp.department_id = d.id JOIN organizations o ON d.organization_id = o.id
        WHERE o.name = :name AND r.probability * r.severity > 4
        ORDER BY r.probability * r.severity DESC
        """)
        sql_count = text("""
                SELECT COUNT(wp.name)
                FROM (result r JOIN dangers_dict ON r.danger_id = dangers_dict.id) 
                JOIN work_places wp ON wp.id = r.work_place_id
                JOIN departments d ON wp.department_id = d.id JOIN organizations o ON d.organization_id = o.id
                WHERE o.name = :name AND r.probability * r.severity > 4
                ORDER BY r.probability * r.severity DESC
                """)
        dic = {'name': org_name}
        return self.conn.execute(sql, **dic), int(self.conn.execute(sql_count, **dic).fetchone()[0])
