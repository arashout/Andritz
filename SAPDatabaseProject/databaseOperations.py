'''
This module contains functions relating to the mySQL database,
especially concerning setup and termination
'''
# TODO
# Implement the mysql connector


def setup_db(db_name):
    '''
    CREATE DATABASE AndritzCoop;
    '''
    command = "CREATE DATABASE {0}".format(db_name)


def grant_permissions(db_name, username, host, password):
    '''
    GRANT ALL PRIVILEGES ON dbTest.* To 'arash'@'localhost' IDENTIFIED BY 'password';
    '''
    command = "GRANT ALL PRIVILEGES ON {0}.* TO '{1}'@'{2}' IDENTIFIED BY {3}".format(
        database_name, username, password)


def select_db(db_name):
    '''
    USE DATABASE AndritzCoop;
    '''
    command = "USE {0}".format(db_name)


def setup_SAP_material_table():
    command = """
    CREATE TABLE sap_materials(
    mat_num varchar(9) NOT NULL,
    mat_type varchar(4) NOT NULL,
    description text,
    basic_mat varchar(255),
    amc varchar(11),
    PRIMARY KEY (mat_num)
    )
    """.replace('\n', '')
    return command

def insert_material(matObj):
