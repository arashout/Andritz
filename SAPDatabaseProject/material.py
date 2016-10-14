class Material(object):
    def __init__(self, list_info):
        '''
        Assuming that list_info is fed from parse_SAP_export
        '''
        self.mat_num = list_info[2]
        self.mat_type = list_info[3]
        self.desc = list_info[5]
        self.basic_mat = list_info[6]
        self.amc = list_info[9].replace(' ','')

    def __str__(self):
        string_rep = """
        {0},{1},{2},{3},{4}
        """.format(
            self.mat_num,
            self.mat_type,
            self.desc,
            self.basic_mat,
            self.amc
        ).replace('\n', '')
        return string_rep

    def db_insert(self, cursor):
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
        print(self.amc)
        cursor.execute(command)


def create_materials_from_SAP_file(file_path):
    with open(file_path, 'r') as f:
        list_rows = f.readlines()

    list_obj = []
    # Skip header rows
    for i in range(1, len(list_rows)):
        row = list_rows[i]
        list_obj.append(Material(row.split('\t')))

    return list_obj


def setup_table(cursor):
    command = """
    CREATE TABLE sap_materials(
    mat_num char(9) NOT NULL,
    mat_type char(4) NOT NULL,
    description text,
    basic_mat varchar(255),
    amc char(11),
    PRIMARY KEY (mat_num)
    )
    """.replace('\n', '')
    cursor.execute(command)
