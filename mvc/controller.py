from .view import View
from .model import Model


class Controller(object):
    def __init__(self, view: View, model: Model):
        self.model = model
        self.view = view

    def add_class(self, classtime):
        self.model.add_class(classtime)

    def check_available_rooms(self):
        pass

    def update_view(self):
        self.view.update(self.model)
