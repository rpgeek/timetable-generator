import datetime

from models.classtime import Classtime
from mvc.excel_writer import ExcelWriter
from mvc.storage import LocalStorage

from .view import View
from .model import Model


class Controller(object):
    def __init__(self, view: View, model: Model):
        self.model = model
        classes = self.model.get_all_classes()
        params = self.model.get_params()

        print('restored classes, params', classes, params)
        self.view = view
        self.view.update_params(params)
        self.view.update_classes(classes)

    def add_class(self, classtime):
        self.model.add_class(classtime)
        self.view.add_classtime(classtime)

    def check_available_rooms(self):
        pass

    def set_params(self, params):
        print(params)
        self.model.set_params(params)


def test_controller_mvc():
    groups = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    groups_division = [2, 3, 2, 3]
    date_start = datetime.date(2017, 10, 2)
    day_max = datetime.date(2018, 2, 22)

    view = ExcelWriter(start=str(date_start), max=str(day_max), groups=groups, division=groups_division)

    model = LocalStorage('db_test.db')
    ctr = Controller(view, model)
    params = {'start': str(date_start), 'max': str(day_max), 'groups': groups, 'division': groups_division}
    ctr.set_params(params)

    today = datetime.date.today()

    cls = Classtime('Biologia', today, 1, 4, [1, 0, 0, 1], 'Mr X', '103E')
    ctr.add_class(cls)

    today += datetime.timedelta(days=1)
    ctr.add_class(Classtime('Historia', today, 1, 4, [1, 1, 1, 1], "aaa", "asd"))
    view.update('out.xlsx')


def test_reload_from_local_storage():
    groups = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    groups_division = [2, 3, 2, 3]
    date_start = datetime.date(2017, 10, 2)
    day_max = datetime.date(2018, 2, 22)

    params = {'start': str(date_start), 'max': str(day_max), 'groups': groups, 'division': groups_division}
    view = ExcelWriter()
    view.update_params(params)

    model = LocalStorage('db_reuse.db')
    ctr = Controller(view, model)
    params = {'start': str(date_start), 'max': str(day_max), 'groups': groups, 'division': groups_division}
    ctr.set_params(params)

    today = datetime.date.today()

    cls = Classtime('Biologia', today, 1, 4, [1, 0, 0, 1], 'Mr X', '103E')
    ctr.add_class(cls)

    today += datetime.timedelta(days=1)
    ctr.add_class(Classtime('Historia', today, 1, 4, [1, 1, 1, 1], "aaa", "asd"))

    view = ExcelWriter()

    model = LocalStorage('db_reuse.db')
    ctr = Controller(view, model)
    view.update('out_reuse.xlsx')
