import datetime

from models.classtime import Classtime
from mvc.excel_writer import ExcelWriter
from mvc.storage import LocalStorage

from .view import View
from .model import Model


class Controller(object):
    def __init__(self, view: View, model: Model):
        self.model = model
        self.view = view

    def add_class(self, classtime):
        self.model.add_class(classtime)
        self.view.add_classtime(classtime)


    def check_available_rooms(self):
        pass



def test_controller_mvc():
    groups = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    groups_division = [2, 3, 2, 3]
    date_start = datetime.date(2017, 10, 2)
    day_max = datetime.date(2018, 2, 22)


    view = ExcelWriter(date_start, day_max, groups, groups_division)

    model = LocalStorage('db_test.db')
    ctr = Controller(view, model)

    today = datetime.date.today()

    cls = Classtime('Biologia', today, 1, 4, [1, 0, 0, 1], 'Mr X', '103E')
    ctr.add_class(cls)
    today += datetime.timedelta(days=1)
    ctr.add_class(Classtime('Historia', today, 1, 4, [1,1,1,1], "aaa", "asd"))
    view.update('out.xlsx')

