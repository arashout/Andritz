import material
import pymysql

# File path to SAP search files
file_path = r"hex nut.txt"
list_material_objects = material.create_materials_from_SAP_file(file_path)


conn = pymysql.connect(host='localhost', port=3306,
                       user='root', passwd='root', db='mysql')
cur = conn.cursor()

cur.execute("USE andritz;")

#material.setup_table(cur)

for mat in list_material_objects:
    mat.db_insert(cur)

cur.close()
conn.close()
