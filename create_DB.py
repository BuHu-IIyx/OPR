from sqlalchemy import create_engine, MetaData, Table, String, Integer, Column, Text, Boolean, ForeignKey


def create_data_base(name: str) -> None:
    metadata = MetaData()
    engine = create_engine(f'sqlite:///{name}.db')
    engine.connect()

    labs = Table('labs', metadata,
                 Column('id', Integer(), primary_key=True),
                 Column('name', String(300), nullable=False),
                 Column('address', String(300), nullable=False),
                 Column('phone', String(20), nullable=False),
                 Column('email', String(40), nullable=False),
                 Column('reg_num', String(20), nullable=False),
                 Column('date_start', String(20), nullable=False),
                 Column('date_off', String(20), nullable=False))

    experts = Table('experts', metadata,
                    Column('id', Integer(), primary_key=True),
                    Column('name', String(100), nullable=False),
                    Column('position', String(100), nullable=False))

    organizations = Table('organizations', metadata,
                          Column('id', Integer(), primary_key=True),
                          Column('name', String(200), nullable=False),
                          Column('head_position', String(100), nullable=False),
                          Column('head_name', String(100), nullable=False))

    departments = Table('departments', metadata,
                        Column('id', Integer(), primary_key=True),
                        Column('name', String(300), nullable=False),
                        Column('address', String(300), nullable=True),
                        Column('organization_id', Integer(), ForeignKey('organizations.id')))

    work_places = Table('work_places', metadata,
                        Column('id', Integer(), primary_key=True),
                        Column('name', String(200), nullable=False),
                        Column('equipment', Text(), nullable=False),
                        Column('department_id', Integer(), ForeignKey('departments.id')))

    dangers_type = Table('dangers_type', metadata,
                         Column('id', Integer(), primary_key=True),
                         Column('name', String(60), nullable=False))

    dangers_dict = Table('dangers_dict', metadata,
                         Column('id', String(10), primary_key=True),
                         Column('description', Text(), nullable=False),
                         Column('type', Integer(), ForeignKey('dangers_type.id')))

    dangers_group = Table('dangers_group', metadata,
                          Column('id', Integer(), primary_key=True),
                          Column('name', String(20), nullable=False),
                          Column('danger_id', String(10), ForeignKey('dangers_dict.id'), nullable=False),
                          Column('probability', Integer(), default=0),
                          Column('severity', Integer(), default=0),
                          Column('comment', Text(), nullable=False),
                          Column('measures', Text(), nullable=True))

    result = Table('result', metadata,
                   Column('id', Integer(), primary_key=True),
                   Column('work_place_id', Integer(), ForeignKey('work_places.id')),
                   Column('danger_id', String(10), ForeignKey('dangers_dict.id')),
                   Column('probability', Integer(), default=0),
                   Column('severity', Integer(), default=0),
                   Column('comment', Text(), nullable=False),
                   Column('measures', Text(), nullable=True))

    metadata.create_all(engine)
