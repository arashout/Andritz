#  -*- coding: utf-8 -*-
import pymysql
import unittest


class CreateTableTestCase(unittest.TestCase):
    conn = pymysql.connect(host='localhost', port=3306,
                           user='arash', passwd='main', db='andritz',
                           use_unicode=True, charset="utf8")
    cur = conn.cursor()

    def setUp(self):
        '''
        '''
        pass

    def test_create_table(self):
        '''
        A command that setups the  test sap_materials table if it doesn't exist
        '''
        command = """
        CREATE TABLE IF NOT EXISTS andritz.test_sap_materials(
        mat_num CHAR(9) NOT NULL,
        mat_type CHAR(4) NOT NULL,
        description BLOB,
        basic_mat VARCHAR(255),
        amc CHAR(11),
        PRIMARY KEY (mat_num)
        )
        """.replace('\n', '')
        success = self.cur.execute(command)
        # If command hits an error return value == E
        assert(success == 0)

    def tearDown(self):
        '''
        Method to delete test table and close all connections
        '''
        command = """
        DROP TABLE IF EXISTS andritz.test_sap_materials
        """
        self.cur.execute(command)
        self.cur.close()
        self.conn.close()


if __name__ == '__main__':
    unittest.main()
