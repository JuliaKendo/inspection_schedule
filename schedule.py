import re
import argparse
import datetime
import random
from collections import defaultdict
from docxtpl import DocxTemplate

ONE_HOUR = 3600


def get_next_time(start_time, end_time, step):
    for i in range(14):
        if start_time > end_time:
            break
        add_to_time = random.randint(ONE_HOUR * 1.5, ONE_HOUR * 2) + (step * 60)
        start_time += datetime.timedelta(seconds=add_to_time)
        yield start_time


def get_check_list(start_time, end_time, step):
    last_check_time = end_time - datetime.timedelta(seconds=ONE_HOUR)
    return [times.strftime('%H:%M %d.%m.%Y') for times in get_next_time(start_time, last_check_time, step)]


def get_schedule_date(string_date, start_hour):
    result = re.findall(r'^(\d{1,2})[-., /](\d{1,2})[-., /](\d{2,4})$', string_date)
    if not result:
        return datetime.datetime.now()
    day, month, year = result[0]
    return datetime.datetime(int(year), int(month), int(day), start_hour, 0, 0)


def create_parser():
    parser = argparse.ArgumentParser(description='Параметры запуска скрипта')
    parser.add_argument('-f', '--file', default='generated_doc.docx', help='Файл ворда с расписанием проверок')
    parser.add_argument('-t', '--template', default='template.docx', help='Шаблон docx')
    parser.add_argument('-d', '--date', help='Дата формирования расписания в формате %%dd.%%mm.%%YYYY', type=str)
    parser.add_argument('-o', '--objects', nargs='+', default=['ГСМ', 'Стоянка', 'Ангар'], help='Список объектов охраны')
    parser.add_argument('-r', '--hour', default=8, help='Час начала расчета расписания', type=int)
    parser.add_argument('-s', '--step', default=5, help='Время между проверками', type=int)
    return parser


def main():
    context = {}
    schedule = defaultdict(list)
    parser = create_parser()
    args = parser.parse_args()

    try:
        schedule_date = get_schedule_date(args.date, args.hour) if args.date else datetime.datetime.now()
        start_time = datetime.datetime(schedule_date.year, schedule_date.month, schedule_date.day, args.hour, 0, 0)
        end_time = start_time + datetime.timedelta(hours=24)

        for num, obj in enumerate(args.objects):
            schedule[f'{num+1}. {obj}'] = get_check_list(start_time, end_time, args.step)

        context['objects'] = schedule.items()
        context['start_date'] = start_time.strftime('%d.%m.%Y')
        context['end_date'] = end_time.strftime('%d.%m.%Y')

        doc = DocxTemplate(args.template)
        doc.render(context)
        doc.save(args.file)

    except (KeyError, TypeError, ValueError) as error:
        print(f'Ошибка создания файла расписания: {error}')

    except OSError as error:
        print(f'Ошибка чтения/записи файла: {error}')

    else:
        print(f'Файл расписания {args.file} успешно создан')


if __name__ == "__main__":
    main()
