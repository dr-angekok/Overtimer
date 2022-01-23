from overtimer import config_crud, xl_parsers, core
import datetime

PERSONAL_FILENAME = 'tests/fixtures/persons.xls'
ORDER_FILENAME = 'tests/fixtures/orders.xlsx'
SETTINGS_INI_PATH = 'tests/fixtures'


def test_core():
    today = datetime.datetime.today()
    config = config_crud.CrudConfig(SETTINGS_INI_PATH)
    test_data = xl_parsers.get_personal_df(PERSONAL_FILENAME)
    number = xl_parsers.get_order_number(ORDER_FILENAME)
    contexts = core.make_contexts(number, test_data, config.chief)
    assert len(contexts) == 7
    for key in contexts[0].keys():
        assert key in ('tabn', 'FIO', 'date', 'year', 'term', 'full_date', 'prnumber', 'start', 'finish', 'minut', 'chief')
    context = contexts[0]
    assert context['tabn'] == 10001
    assert context['FIO'] == 'Бидон Надоев'
    assert context['term'] == 2
    assert context['prnumber'] == '278/ка'
    assert context['year'] == today.strftime('%Y')[2:]
    assert context['full_date'] == today.strftime('%d.%m.%Y')
