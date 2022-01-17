from overtimer import xl_parsers

PERSONAL_FILENAME = 'tests/fixtures/persons.xls'


def test_personal_parser():
    test_data = xl_parsers.personal_parser(PERSONAL_FILENAME)
    assert test_data.shape == (7, 3)
    for item in test_data.columns:
        assert item in ['names', 'numbers', 'hours']
    assert 'Бидон Надоев' in test_data['names'].values
    assert 28448 in test_data['numbers'].values
    assert 2 in test_data['hours'].values
