'''
This module contains functions relating to the mySQL database,
especially concerning setup and termination
'''

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