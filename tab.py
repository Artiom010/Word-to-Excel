import sqlite3

def show_tables():
    conn = sqlite3.connect('facturi.db')
    cursor = conn.cursor()

    # Selectează tabelele existente
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()

    # Afișează tabelele
    print("Tabelele existente:")
    for table in tables:
        print(table[0])

    conn.close()

show_tables()
