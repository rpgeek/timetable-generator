from tinydb import TinyDB


class LocalDatabase():
    def __init__(self):
        super().__init__()

    def connect(self, name):
        try:
            self.__db = TinyDB(name)
            self.__connected = True
            self.days_tbl = self.__db.table('days')
        except Exception as e:
            print("error in connection", e)
            self.__connected = False

    def find_day_by_date(self, date):
        print('find day by date')
        day = self.days_tbl.search('day' == str(date))
        print(day)
        return day

    def add_new_day(self, day, name, resp, hr_st, hr_end, groups, room):

        day_dict = dict(day=str(day), name=name, responsible=resp,
                        hours_start=hr_st, hours_end=hr_end,
                        groups=groups, room=room)

        self.days_tbl.insert(day_dict)

    def get_day(self, id):
        self.days_tbl.find_one(id=id)

    def update_day(self, day_dict, what):
        self.days_tbl.update(day_dict, [what])

    def delete_day(self, id):
        self.days_tbl.delete(day=id)
