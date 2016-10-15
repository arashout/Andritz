from csv import reader
import pymysql

class Material(object):
    def __init__(self, list_info):
        '''
        Assuming that list_info is fed from parse_SAP_export
        '''
        self.mat_num = list_info[2]
        self.mat_type = list_info[3]
        self.desc = list_info[5]
        self.basic_mat = list_info[6]
        self.amc = list_info[9].replace(' ', '')

    def __str__(self):
        string_rep = """
        {0}\t{1}\t{2}\t{3}\t{4}
        """.format(
            self.mat_num,
            self.mat_type,
            self.desc,
            self.basic_mat,
            self.amc
        ).replace('\n', '')
        return string_rep

    def db_insert(self, cursor):
        try:
            command = """
            INSERT INTO sap_materials (
            mat_num,mat_type,description,basic_mat,amc
            )
            VALUES ('{0}','{1}','{2}','{3}','{4}')
            """.format(
                self.mat_num,
                self.mat_type,
                self.desc,
                self.basic_mat,
                self.amc
            ).replace('\n', '')
            cursor.execute(command)
        except pymysql.err.ProgrammingError as e:
            print("Caught a Programming Error:")
            print(e)
        except pymysql.err.IntegrityError as e:
            print("Caught a Integrity Error:")
            print(e)


def create_materials_from_SAP_file(file_path):
    list_obj = []
    count = 0
    with open(file_path, 'r') as f:
        # CSV module doing most of the heavy lifting here
        for line in reader(f):
            # Skip the header
            if count == 0:
                pass
            else:
                list_obj.append(Material(line))
            count = + 1
    return list_obj


def setup_table(cursor):
    command = """
    CREATE TABLE IF NOT EXISTS sap_materials(
    mat_num char(9) NOT NULL,
    mat_type char(4) NOT NULL,
    description text,
    basic_mat varchar(255),
    amc char(11),
    PRIMARY KEY (mat_num)
    )
    """.replace('\n', '')
    cursor.execute(command)
