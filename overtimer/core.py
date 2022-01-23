import datetime

NORMAL_DAY = (16, 0)
HALF_DAY = (14, 45)


def number_to_month(number):
    month = (
        'января',
        'февраля',
        'марта',
        'апреля',
        'мая',
        'июня',
        'июля',
        'августа',
        'сентября',
        'октября',
        'ноября',
        'декабря'
    )
    return month[int(number) - 1]


def make_contexts(order_number, persons_data, chief):
    today = datetime.datetime.today()
    day = today.strftime("%d")
    month = today.strftime('%m')
    year = today.strftime('%Y')
    full_date = today.strftime('%d.%m.%Y')

    weekday = today.weekday()
    start, minutes = HALF_DAY if weekday == 4 else NORMAL_DAY

    contexts = []
    for _, row in persons_data.iterrows():
        context = {
            'tabn': row['numbers'],
            'FIO': row['names'],
            'date': '{}. {}'.format(day, number_to_month(month)),
            'year': year[2:],
            'term': row['hours'],
            'full_date': full_date,
            'prnumber': order_number,
            'start': start,
            'finish': start + row['hours'],
            'minut': minutes,
            'chief': chief
        }
        contexts.append(context)
    return contexts
