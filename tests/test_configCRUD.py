from overtimer import config_crud
import os


SETTINGS_INI_PATH = 'tests/fixtures'


def test_init(tmpdir):
    config = config_crud.CrudConfig(tmpdir)
    testpath = os.path.join(tmpdir, 'settings.ini')
    assert config.path == testpath
    assert os.path.exists(testpath)
    assert config.template_path == '/'
    assert config.person_list_path == '/'
    assert config.orders_path == '/'
    assert config.chief == 'ФИО'


def test_r_w(tmpdir):
    config = config_crud.CrudConfig(tmpdir)
    config.template_path_set('/template.wrld')
    config.person_list_path_set('/person_list.xlsx')
    config.orders_path_set('/orders.xlsx')
    config = config_crud.CrudConfig(tmpdir)
    assert config.template_path == '/template.wrld'
    assert config.person_list_path == '/person_list.xlsx'
    assert config.orders_path == '/orders.xlsx'
    assert config.chief == 'ФИО'


def test_load():
    config = config_crud.CrudConfig(SETTINGS_INI_PATH)
    assert config.template_path == 'tests/fixtures/template.docx'
    assert config.person_list_path == 'tests/fixtures/persons.xlsx'
    assert config.orders_path == 'tests/fixtures/orders.xlsx'
    assert config.chief == 'ФИО'

