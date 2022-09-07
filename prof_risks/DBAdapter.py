from sqlalchemy import create_engine


class DBAdapter:
    def __init__(self, name):
        self.engine = create_engine(f'sqlite:///{name}.db')

    def add_organization(self):
        pass
