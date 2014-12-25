# coding: utf-8

from xlutils.filter import BaseFilter, GlobReader, process, DirectoryWriter

__author__ = 'valiullin'


class ValueRowFilter(BaseFilter):

    goodlist = None

    def __init__(self):
        self.wtrowx = -1
        self.skipped_row = -1

    def workbook(self, rdbook, wtbook_name):
        self.next.workbook(rdbook, 'filtered_' + wtbook_name)

    def row(self, rdrowx, rwrowx):
        value = self.rdsheet.cell(rdrowx, ROW).value
        if value == VALUE:
            self.skipped_row = rdrowx
            self.wtrowx += 1

    def cell(self, rdrowx, rdcolx, wtrowx, wtcolx):
        if self.skipped_row == rdrowx:
            self.next.cell(rdrowx, rdcolx, self.wtrowx, wtcolx)


xls_file = 's.xls'
results_folder = ''

VALUE = 'XXX'
ROW = 1

process(
    GlobReader(xls_file),
    ValueRowFilter(),
    DirectoryWriter(results_folder)
)