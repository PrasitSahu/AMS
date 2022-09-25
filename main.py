import mysql.connector as mysql
from mysql.connector import Error
from env import password
from datetime import date
import os
from os.path import exists
from openpyxl import Workbook


def cls():
    """ clears the console's screen """

    os.system('cls' if os.name == 'nt' else 'clear')


months = ['january', 'february', 'march', 'april', 'may', 'june',
          'july', 'august', 'september', 'october', 'november', 'december']


def record_attendance():
    """ Adds new records to the attendance table/register """

    # get today's date
    today = date.today()
    this_year = today.year
    this_month = today.month
    this_day = today.day

    # connection to students main database
    connection = connect_db('classxii')
    cursor = connection.cursor()
    table = f"{months[this_month]}_{this_year}"

    # gets all the tables from database
    cmd = f"SHOW TABLES;"
    cursor.execute(cmd)
    tables = []
    for i in cursor:
        tables.append(i[0])

    if table not in tables:
        cmd = f"CREATE TABLE {table}(" \
              f"id int NOT NULL PRIMARY KEY," \
              f"name varchar(60) NOT NULL" \
              f");"
        # add tables
        cursor.execute(cmd)
        connection.commit()

        # add students data to the new table
        cmd = f"INSERT INTO {table}(id, name) SELECT id, name FROM attendance;"
        cursor.execute(cmd)
        connection.commit()

    cursor.execute(f'SELECT * FROM {table};')
    # stores all the date columns from the database
    prev_dates = cursor.column_names[2:]
    col = f"{this_day}d_{this_month}m_{this_year}"

    if col not in prev_dates:
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
        cmd = f"ALTER TABLE {table} ADD {col} varchar(7)"

        # add new date column
        cursor.execute(cmd)
        connection.commit()

        try:
            for i in range(len(record)):
                cmd = f"UPDATE {table} SET {col} = '{record[i]}' WHERE id = {i + 1};"
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

        except Error as error:
            # update failed message as an error
            print(f"Record Update Failed !: {error}")

            # reverting changes because of exception
            connection.rollback()
    else:
        print('Error: Attendance record for this date is already available!\n')


def update_records():
    wb = Workbook()
    ws = wb.active

    # connection to students main database
    connection = connect_db('classxii')
    cursor = connection.cursor()

    cmd = "SHOW TABLES;"
    cursor.execute(cmd)
    cls()
    print("select a table - ")
    num = 1
    tables = []
    for table in cursor:
        if table[0].lower() == 'attendance':
            continue
        print(f"{num}. {table[0]}")
        tables.append(table[0])
        num += 1
    table_num = None
    while True:
        try:
            table_num = int(input('\n> '))
            if table_num in range(1, len(tables) + 1):
                break
            else:
                print('Invalid input, try again...')
                continue
        except ValueError:
            print('Invalid input, try again...')
            continue
    cursor.execute(f'SELECT * FROM {tables[table_num - 1]};')

    # stores all the columns from the database
    columns = cursor.column_names[:]
    row = 1
    col = 0
    cols = ['A', 'B', 'C', 'D', 'E', '  F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N',
            'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE']
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

    file = f'{tables[table_num - 1]}.xlsx'
    if exists(file):
        os.remove(file)
    wb.save(file)
    print('updated successfully. :)\n')


def connect_db(db_name):
    """ makes a connection to mysql db """

    connection = mysql.connect(host='ams-db.cw84lzcpsrau.ap-south-1.rds.amazonaws.com',
                               user='vicky',
                               password=password or '',
                               database=db_name)
    return connection


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

    opt = None
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
            It is a very useful tool for maintaining an attendance record. The interface is very easy to understand.
            To mark as 'Present' just type 'Yes' or 'Y' and to mark as 'Absent' just type 'No' or 'N'. 
                      You can even save the records as an excel sheet by choosing the 'UPDATE IN EXCEL SHEET' option
            from the main menu. 
        ''')
        input('go back > ')
        cls()
        main()


if __name__ == '__main__':
    main()
