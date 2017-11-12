import os
from tinydb import TinyDB
from .model import Model
from .model import Classtime

class LocalStorage(Model):
    def __init__(self, name='db.json'):
        super().__init__()
        self.__connect(name)

    def __connect(self, name):
        try:
            self.__db = TinyDB(name)
            self.__connected = True
            self.days_tbl = self.__db.table("days")
            self.params_tbl = self.__db.table("params")
        except Exception as e:
            print("error in connection", e)
            self.__connected = False

    def __find_day_by_date(self, date):
        print('find day by date')
        day = self.days_tbl.search('day' == str(date))
        print(day)
        return day

    def add_class(self, classtime: Classtime):
        self.days_tbl.insert(classtime.to_dict)

    def get_all_classes(self):
        return self.days_tbl.all()

    def remove_class(self, ids):
        self.days_tbl.remove(doc_ids=ids)

    def set_params(self, params):
        if len(self.params_tbl.all()) == 0:
            self.params_tbl.insert(params)
        else:
            print("db has params, set_params")

    def get_params(self):
        params = self.params_tbl.all()
        return params[0]




def test_answer():
    db = LocalStorage('db.json')
    assert db != None


def remove(param):
    os.remove(param)

def api_test():
    db = LocalStorage('db.json')

    from datetime import date
    today = date.today()
    cls = Classtime('Biologia', today, 0, 4, [1, 2, 3], 'Mr X', '103E')
    db.add_class(cls)
    assert len(db.get_all_classes()) == 1

    id = db.get_all_classes()[0].doc_id

    db.remove_class([id])
    assert len(db.get_all_classes()) == 0

    remove('db.json')


