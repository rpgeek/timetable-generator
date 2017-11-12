from models.classtime import Classtime

class Model(object):

    def add_class(self, classtime: Classtime):
        raise NotImplementedError()

    def remove_class(self, classtime: Classtime):
        raise NotImplementedError()

    def get_all_classes(self):
        raise NotImplementedError()

    def set_params(self, params):
        raise NotImplementedError()

    def get_params(self):
        raise NotImplementedError()
