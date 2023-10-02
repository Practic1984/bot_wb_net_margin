import sqlite3
from operator import itemgetter
import pandas as pd
from datetime import datetime


db_name = 'users.db'

def add_users(message, status):
    print(message)
    # reg_time = datetime.fromtimestamp(message.date)
    # print(reg_time)
    # reg_time = reg_time.strftime('%Y-%m-%d %H:%M:%S')
    # print(reg_time)
    connect = sqlite3.connect(db_name)
    cursor = connect.cursor()

    cursor.execute("""CREATE TABLE IF NOT EXISTS users(
        from_user_id TEXT,
        username TEXT,
        first_name TEXT,
        reg_date INTEGER, 
        date_str TEXT,
        status TEXT,
        mail TEXT 
)
     """)
    connect.commit()

    perem = message.from_user.id

    cursor.execute('SELECT from_user_id FROM users WHERE from_user_id = ?', [perem])
    data_test = cursor.fetchone()

    data_discont ={}
    result = ''
    if data_test is None:

        reg_time = datetime.fromtimestamp(message.date)
        reg_time = reg_time.strftime('%Y-%m-%d %H:%M:%S')
        cursor.execute('INSERT INTO users VALUES (?,?,?,?,?,?,?)', [message.from_user.id, message.from_user.username, message.from_user.first_name, message.date, reg_time, status, 'mail'])
        connect.commit()
        cursor.close()

        connect.commit()
        cursor.close()

def del_table(name_db):
    connect = sqlite3.connect(name_db)
    cursor = connect.cursor()
    cursor.execute("DROP TABLE IF EXISTS wb_item")
    cursor.close()
    return True

def search_db(key):
    connect = sqlite3.connect(key)
    cursor = connect.cursor()
    df = pd.read_sql_query("SELECT * FROM wb_item", connect)
    df.to_excel('all_base.xlsx', index=False)
    cursor.close()












