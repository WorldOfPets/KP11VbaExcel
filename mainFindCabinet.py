import openpyxl
import sys
from basedef import get_lesson, get_day, get_lesson_num
from dbfunc import get_cab_db


days = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота"]

def find_cabinet(filePath, row_id):
    wb = openpyxl.load_workbook(filePath, read_only=False, keep_vba=True)

    ws = wb.active
    lessons = get_lesson(ws)

    cabinets = get_cab_db()
    print(cabinets)
    day = get_day(ws, 1, int(row_id))
    less_num = get_lesson_num(ws, 2, int(row_id))
    for less in lessons:
        if less.lesson_day == day and str(less.lesson_num) == str(less_num):
            if str(less.cabinet) in cabinets:
                try:
                    cabinets.remove(str(less.cabinet))
                except Exception as ex:
                    print(".", end="")
    print(f"Вот список свободных кабинетов на {day} на {less_num} урок:")
    print(cabinets)


def main():
    if sys.argv is not None and len(sys.argv) > 1:
        filePath = sys.argv[1]
        find_cabinet(filePath, sys.argv[2])

if __name__ == '__main__':
    main()
x = input()