#  -*- coding: utf-8 -*-
import material
import pymysql

# File path to SAP search files
file_path = r"hex nut.csv"
list_material_objects = material.create_materials_from_SAP_file(file_path)


conn = pymysql.connect(host='localhost', port=3306,
                       user='root', passwd='root', db='mysql')
cur = conn.cursor()

cur.execute('SET NAMES utf8mb4;')
cur.execute('SET CHARACTER SET utf8mb4;')
cur.execute('SET character_set_connection=utf8mb4;')

cur.execute("USE andritz;")

for mat in list_material_objects:
    mat.db_insert(cur)

conn.commit()
cur.close()
conn.close()
