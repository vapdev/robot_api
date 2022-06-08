from lib2to3.pgen2.pgen import DFAState
from sqlalchemy import create_engine
import pandas as pd
import cx_Oracle

class dbConnection():
    
    def __init__(self):
        self.engine = create_engine('oracle://desenv:max@192.168.180.3:1521/desv')

    def query(self, query):
        df = pd.read_sql(query, self.engine)
        return df
