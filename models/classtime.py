
class Classtime(object):
    def __init__(self, name, date, hrs_st,
                 hrs_stp, group, lead_name):
        self.lead_name = lead_name
        self.gropus = group
        self.hrs_stp = hrs_stp
        self.hrs_st = hrs_st
        self.date = date
        self.name = name
