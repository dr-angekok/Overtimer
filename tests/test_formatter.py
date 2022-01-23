from overtimer.formaters import make_docx
from os import listdir, path
import datetime


def test_make_docx(tmpdir):
    context = {
            'tabn': 12345,
            'FIO': 'Василий Пупновский',
            'date': '01. января',
            'year': 22,
            'term': 4,
            'full_date': '01.01.1900',
            'prnumber': '123/ka',
            'start': '16',
            'finish': '18',
            'minut': '00',
    }
    contexts = (context for _ in range(3))
    today = datetime.datetime.today()
    filename = '{}.docx'.format(today.strftime("%d%b%Y"))
    make_docx(contexts, tmpdir)
    assert path.isfile(path.join(tmpdir, filename))
    count = len([f for f in listdir(tmpdir) if path.isfile(path.join(tmpdir, f))])
    assert count == 1
