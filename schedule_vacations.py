#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import datetime
import locale
import pandas as pd
import pathlib
import random
import sys

DELIMITER_1 = '-'*50
DELIMITER_2 = '\n'+'='*52
choices = {}
not_allocated = []

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('staff', type=str,\
        help='Файл с фамилиями сотрудников')
    parser.add_argument('-s', type=str, default='2020-01-01',\
        help='Дата начала отпусков')
    parser.add_argument('-e', type=str, default='2020-12-31',\
        help='Дата конца отпусков')
    parser.add_argument('-f', type=str, default='4W-MON',\
        help='Формат интервала отпуска')
    parser.add_argument('-i1', type=int, default=0,\
        help='Индекс начала первого интервала')
    parser.add_argument('-i2', type=int, default=-1,\
        help='Индекс конца первого интервала')
    options = parser.parse_args()
    staff_file_name = options.staff
    start_date = options.s
    end_date = options.e
    frequency = options.f
    i1 = options.i1
    i2 = options.i2

    locale.setlocale(locale.LC_ALL, '')
    date_time = datetime.datetime.now()
    print('Время: {}\n'.format(date_time.strftime('%d %B %Y, %H:%M:%S')))

    staff = pd.read_csv(staff_file_name, header=None, names=['Сотрудник'])
    staff = staff['Сотрудник']
    print('Сотрудники'.upper())
    print('Перечень:\n {}'.format(', '.join(staff)))
    print('Количество: {}'.format(len(staff)))
    random.shuffle(staff)
    print('Случайная очерёдность выбора:\n {}\n'.format(', '.join(staff)))

    print('Параметры'.upper())
    print('s={}, e={}, f={}, i1={}, i2={}\n'.format(
        start_date, end_date, frequency, i1, i2))

    print('Интервалы времени'.upper())
    date_range = pd.date_range(start=start_date, end=end_date, freq=frequency)
    print('Количество: {}'.format(len(date_range)))
    if len(date_range) < len(staff):
        print('Количество интервалов времени меньше количества сотрудников!')
        print('Измените значения параметров запуска.')
        sys.exit()

    parts_generator = range(1, len(date_range)+1)
    all_parts = [i for i in parts_generator]
    print('Порядковые номера: {}'.format(all_parts))
    if i1 == 0 and i2 == -1:
        parts = [all_parts,]
    else:
        parts = [all_parts[i1:i2], all_parts[:i1] + all_parts[i2:]]
    for number, part in enumerate(parts):
        print('Часть {}: {}'.format(number+1, part))
        if len(part) < len(staff):
            print('Количество вариантов выбора меньше количества сотрудников!')
            print('Измените значения параметров запуска.')
            sys.exit()

    for worker in staff:
        choices[worker] = []

    for part in parts:
        print(DELIMITER_2)
        print('Все варианты: {}'.format(part))
        for number, worker in enumerate(staff):
            print('{} {}'.format(number+1, DELIMITER_1))
            print('Сотрудник: {}'.format(worker))
            print('Варианты: {}'.format(part))
            choice = random.choice(part)
            print('Выбор: {}'.format(choice))
            index = part.index(choice)
            del part[index]
            choices[worker].append(choice)
        not_allocated.append(part)

    print(DELIMITER_2)
    print('Результаты'.upper())

    ordinal = pd.DataFrame(choices)
    ordinal = ordinal.T
    print(ordinal)
    print()
    print('Не распределено: {}\n'.format(not_allocated))

    dates = ordinal.stack().sort_values()
    dates = dates.replace(to_replace=parts_generator, value=date_range.strftime('%d %B')) 
    dates.name = 'Дата начала'
    print(dates)
    print()

    export_path = pathlib.Path(date_time.strftime('%Y-%m-%d, %H_%M_%S')).with_suffix('.xlsx')
    dates.to_excel(export_path)
    print('Сохранено в файл: {}'.format(export_path.absolute()))
