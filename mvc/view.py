from models.classtime import Classtime

from .model import Model

class View(object):
    def add_classtime(self, cls: Classtime, repeated=False):
        raise NotImplementedError()

    def update(self, params):
        raise NotImplementedError()

