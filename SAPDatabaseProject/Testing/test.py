#  -*- coding: utf-8 -*-
import pymysql
import unittest


class BaseTestCase(unittest.TestCase):
    @classmethod
    def setUp(self):
        '''
        Establish connections and create the test table
        '''
        self.conn = pymysql.connect(host='localhost', port=3306,
                                    user='arash', passwd='main', db='andritz',
                                    use_unicode=True, charset="utf8")
        self.cur = self.conn.cursor()
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
        # If command hits an error, cursor returns value == E
        self.assertTrue(success == 0, "mysql error")

    @classmethod
    def tearDown(self):
        '''
        Method to delete test table and close all connections
        '''
        self.cur.execute("USE andritz")
        self.cur.execute("DROP TABLE test_sap_materials")


class TableCreationTestCase(BaseTestCase):
    def test_table_exists(self):
        '''
        Check if the table exists
        '''
        command = """
        SELECT *
        FROM information_schema.tables
        WHERE table_schema = 'andritz'
        AND table_name = 'test_sap_materials'
        LIMIT 1;
        """
        self.cur.execute(command)

        self.assertIsNotNone(self.cur.fetchone())


if __name__ == '__main__':
    unittest.main()
