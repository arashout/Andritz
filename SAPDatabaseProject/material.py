from csv import reader
import pymysql


class Material(object):
    def __init__(self, mat_num, mat_type="", desc="", basic_mat="", amc=""):
        '''
        Assuming that list_info is fed from parse_SAP_export
        '''
        self.mat_num = mat_num
        self.mat_type = mat_type
        # mysql doesn't like single quotations, have to escape them
        self.desc = desc
        self.basic_mat = basic_mat
        self.amc = amc

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
            INSERT INTO andritz.sap_materials (
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
            print(e)
        except pymysql.err.IntegrityError as e:
            # If the entry already exists then update it
            # Check if the error message contains key phrase
            if "Duplicate entry" in e.args[1]:
                command = """
                UPDATE andritz.sap_materials
                SET mat_type='{1}',description='{2}',basic_mat='{3}',amc='{4}'
                WHERE mat_num='{0}';
                """.format(
                    self.mat_num,
                    self.mat_type,
                    self.desc,
                    self.basic_mat,
                    self.amc
                ).replace('\n', '')
                cursor.execute(command)
            else:
                print(e)


def create_materials_from_SAP_file(file_path):
    list_obj = []
    count = 0
    with open(file_path, 'r') as f:
        # CSV module doing most of the heavy lifting here
        for row_info in reader(f):
            # Skip the header
            if count == 0:
                pass
            else:
                mat_num = row_info[2]
                mat_type = row_info[3]
                # mysql doesn't like single quotations, have to escape them
                desc = row_info[5].replace("'", "\\'")
                basic_mat = row_info[6]
                amc = row_info[9].replace(' ', '')
                list_obj.append(
                    Material(mat_num, mat_type, desc, basic_mat, amc))
            count = + 1
    return list_obj


def setup_table(cursor):
    command = """
    CREATE TABLE IF NOT EXISTS andritz.sap_materials(
    mat_num CHAR(9) NOT NULL,
    mat_type CHAR(4) NOT NULL,
    description BLOB,
    basic_mat VARCHAR(255),
    amc CHAR(11),
    PRIMARY KEY (mat_num)
    )
    """.replace('\n', '')
    cursor.execute(command)
