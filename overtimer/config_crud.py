import configparser
import os
from pathlib import Path


class CrudConfig:
    def __init__(self, path=os.getcwd()):
        self.config = configparser.ConfigParser()
        self.path = os.path.join(path, 'settings.ini')
        if not os.path.exists(self.path):
            self.create_config()
        self.config.read(self.path)

    def config_save(self):
        if not os.path.exists(self.path):
            file = Path(self.path)
            file.touch(exist_ok=True)
        with open(self.path, "w") as config_file:
            self.config.write(config_file)

    def create_config(self):
        self.config.add_section('paths')
        self.config.add_section('superiors')
        self.config.set('superiors', 'Chief', 'ФИО')
        self.config_save()
        self.orders_path_set('/')
        self.template_path_set('/')
        self.person_list_path_set('/')

    def orders_path_set(self, path):
        self.config.set('paths', 'orders', path)
        self.config_save()

    def template_path_set(self, path):
        self.config.set('paths', 'template', path)
        self.config_save()

    def person_list_path_set(self, path):
        self.config.set('paths', 'person_list', path)
        self.config_save()

    @property
    def orders_path(self):
        return self.config.get('paths', 'orders', fallback='/')

    @property
    def template_path(self):
        return self.config.get('paths', 'template', fallback='/')

    @property
    def person_list_path(self):
        return self.config.get('paths', 'person_list', fallback='/')

    @property
    def chief(self):
        return self.config.get('superiors', 'Chief', fallback='ФИО')
