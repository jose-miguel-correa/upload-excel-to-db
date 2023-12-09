import pymssql
from dotenv import load_dotenv
import os

load_dotenv()

servidor = os.getenv("SQL_SERVER")
database = os.getenv("SQL_DATABASE")
usuario = os.getenv("SQL_USERNAME")
passwd = os.getenv("SQL_USER_PASSWORD")

conn = pymssql.connect(server=servidor, user=usuario, password=passwd, database=database, charset='UTF-8')

cursor = conn.cursor()
