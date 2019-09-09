#!/usr/bin/env python3

import unittest

import exploder

import openpyxl


class TestExploder(unittest.TestCase):

    def test_test1(self):

        wb = openpyxl.load_workbook(filename='test1.xlsx')

        ws1 = wb['Sheet1']

        try:
            ws2 = wb['Sheet2']
        except KeyError:
            ws2 = wb.copy_worksheet(ws1)
            ws2.title = 'Sheet2'

        exploder.explode(wb, ws1, ws2, [1, 3])

        self.assertEqual(ws2['A1'].value, 'foo')
        self.assertEqual(ws2['B1'].value, 'Hello world')
        self.assertEqual(ws2['C1'].value, '1')
        self.assertEqual(ws2['A2'].value, 'foo')
        self.assertEqual(ws2['B2'].value, 'Hello world')
        self.assertEqual(ws2['C2'].value, '2')
        self.assertEqual(ws2['A3'].value, 'foo')
        self.assertEqual(ws2['B3'].value, 'Hello world')
        self.assertEqual(ws2['C3'].value, '3')
        self.assertEqual(ws2['A4'].value, 'bar')
        self.assertEqual(ws2['B4'].value, 'Hello world')
        self.assertEqual(ws2['C4'].value, '1')
        self.assertEqual(ws2['A5'].value, 'bar')
        self.assertEqual(ws2['B5'].value, 'Hello world')
        self.assertEqual(ws2['C5'].value, '2')
        self.assertEqual(ws2['A6'].value, 'bar')
        self.assertEqual(ws2['B6'].value, 'Hello world')
        self.assertEqual(ws2['C6'].value, '3')
        self.assertEqual(ws2['A7'].value, 'yellow')
        self.assertEqual(ws2['B7'].value, 'Goodbye')
        self.assertEqual(ws2['C7'].value, 'two')
        self.assertEqual(ws2['A8'].value, 'yellow')
        self.assertEqual(ws2['B8'].value, 'Goodbye')
        self.assertEqual(ws2['C8'].value, 'three')


if __name__ == '__main__':
    unittest.main()
