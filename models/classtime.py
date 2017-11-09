
class Classtime(object):
    def __init__(self, name, date, hrs_st,
                 hrs_stp, group, lead_name,
                 room):
        self.room = room
        self.lead_name = lead_name
        self.groups = group
        self.hrs_stp = hrs_stp
        self.hrs_st = hrs_st
        self.date = date
        self.name = name

    @property
    def to_dict(self):
        return dict(day=str(self.date), name=self.name, responsible=self.lead_name,
                        hours_start=self.hrs_st, hours_end=self.hrs_stp,
                        groups=self.groups, room=self.room)

    @property
    def text(self):
        return "{} {} {}".format(self.name, self.lead_name, self.room)