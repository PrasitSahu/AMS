import mysql.connector as mysql
from mysql.connector import Error
from mysql.connector import errorcode
from env import password
from datetime import date
import os
from openpyxl import Workbook


def cls():
    """ clears the screen """

    os.system('cls' if os.name == 'nt' else 'clear')


def record_attendance():
    """ Adds new records to the attendance register """

    # connection to students main database
    connection = mysql.connect(host='localhost',
                               user='root',
                               password=password or '',
                               database='student_data')

    cursor = connection.cursor()

    cursor.execute('SELECT * FROM attendance_record;')
    # stores all the date columns from the database
    prev_dates = cursor.column_names[3:]

    # get today's date
    today = date.today()
    this_year = today.year
    this_month = today.month
    this_day = today.day
    col = f"{this_day}d_{this_month}m_{this_year}"

    while True:
        for_today = input('Do you want to record today\'s attendance? ')
        if for_today.lower() in ['y', 'yes']:
            if col in prev_dates:
                print('Warn: The attendance of this date already exists in the database. Please choose any other date.\n')
                continue
            else:
                break
        elif for_today.lower() in ['n', 'no']:
            record_date = input('Enter the date of attendance record in format [dd-mm-yyyy]: ')
            record_date = record_date.split('-')
            col = f"{record_date[0]}d_{record_date[1]}m_{record_date[2]}"
            this_day = record_date[0]
            this_month = record_date[1]
            this_year = record_date[2]
            # d = check_date_format(record_date)
            if col in prev_dates:
                print('Warn: The attendance of this date already exists in the database. Please choose any other date.\n')
                continue
            else:
                break
        else:
            print('Invalid input! Try again...')
            continue

    cls()
    record = []
    print(f'''Record for: {this_day}/ {this_month}/ {this_year}\n''')

    for student in cursor:
        while True:
            prompt = input(f"{student[0]}. {student[1]} - ")
            if prompt.lower() in ['y', 'yes']:
                record.append('present')
                break
            elif prompt.lower() in ['n', 'no']:
                record.append('absent')
                break
            else:
                print('\nInvalid Input, Try again...')
                continue

    # sql command to add data to database
    cmd = f"ALTER TABLE attendance_record ADD {col} varchar(7)"

    # add new date column
    cursor.execute(cmd)
    connection.commit()

    try:
        for i in range(len(record)):
            cmd = f"UPDATE attendance_record SET {col} = '{record[i]}' WHERE rollno = {i + 1};"
            cursor.execute(cmd)
            connection.commit()

        print('\nRecord Updated Successfully. :) \n')

        while True:
            to_update_excel = input('want to update in excel sheet? ')
            if to_update_excel.lower() in ['yes', 'y']:
                update_records()
                break
            elif to_update_excel.lower() in ['no', 'n']:
                break
            else:
                print('\nInvalid input, Try again...')
                continue

    except mysql.connector.Error as error:
        # update failed message as an error
        print("Record Update Failed !: {}".format(error))

        # reverting changes because of exception
        connection.rollback()


def update_records():
    wb = Workbook()
    ws = wb.active

    # connection to students main database
    connection = mysql.connect(host='localhost',
                               user='root',
                               password=password or '',
                               database='student_data')

    cursor = connection.cursor()
    cursor.execute('SELECT * FROM attendance_record;')
    # stores all the columns from the database
    columns = cursor.column_names[:]
    row = 1
    col = 0
    cols = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    for column in columns:
        ws[f"{cols[col]}{row}"] = column
        col += 1

    row = 2
    col = 0
    for student in cursor:
        for info in student:
            ws[f"{cols[col]}{row}"] = info
            col += 1
        row += 1
        col = 0
    os.remove('records.xlsx')
    wb.save('records.xlsx')
    print('updated successfully. :)\n')


def main():
    """ Main (Root) Function """

    cls()
    print('''
                                        ATTENDANCE MANAGEMENT SYSTEM - AMS

        SELECT AN OPTION - 
    ────────────────────────
    1. RECORD ATTENDANCE
    2. UPDATE IN EXCEL SHEET
    3. HOW TO ??

        ''')

    while True:
        try:
            opt = int(input('> '))
            if opt not in [1, 2, 3]:
                print('Invalid input! try again...')
                continue
            break
        except ValueError:
            print('Invalid input! try again...')
            continue

    cls()

    if opt == 1:
        record_attendance()
    elif opt == 2:
        update_records()
    else:
        print('''
                                                    HOW TO USE
                                                   ────────────
        ''')


if __name__ == '__main__':
    main()
