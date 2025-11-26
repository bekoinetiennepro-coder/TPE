import pymysql

def get_connection():
    return pymysql.connect(
        host="localhost",
        user="root",
        password="",   # Mets ton mot de passe MySQL ici
        database="compta",
        charset="utf8mb4",
        cursorclass=pymysql.cursors.DictCursor
    )
