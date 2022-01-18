from overtimer.formaters import make_docx
from os import path


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
    make_docx(contexts, 'text.docx', tmpdir)
    assert path.isfile(path.join(tmpdir, 'text.docx'))
