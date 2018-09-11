import copy
import os

import  openpyxl
from typing import Container

TEMPLATE_DIRECTORY = r'F:\LabData\Lab\ISO 17025\Control Chart\Templates\fastmake ccs'
CHART_DIRECTORY = r'F:\LabData\Lab\ISO 17025\Control Chart\Templates\fastmake ccs\chart copies'

class Chart:
    def __init__(self, file_path, name=None):
        self.file_path = file_path
        self.name = name
        self.wb = None
        self.active_worksheet = None

    def load_workbook(self):
        self.wb = openpyxl.load_workbook(self.file_path)

    def set_active_worksheet(self, target):
        self.active_worksheet = wb.sheets[target]

    def get_cell_copy(self, target):
        return copy.deepcopy(self.active_worksheet[f'{target}'])

    def set_cell(self, target, value):
        self.active_worksheet[f'{target}'] = value

class ChartGatherer:
    def __init__(self, directory, chart_format=Chart):
        self.directory = directory
        self.gathered_charts = []
        self.chart_format = chart_format

    def gather_charts(self, extension='xlsx'):
        for dirpath, dirnames, filenames in os.walk(self.directory):
            filenames = [chart_format(os.path.join(dirpath, file)) for file in filenames if file.endswith(extension)]
            self.gathered_charts.extend(filenames)

class Template:
    def __init__(self, type):
        self.type = type


if name == '__main__':
    chart_gatherer = ChartGatherer(CHART_DIRECTORY)
    chart_gatherer.gather_charts()

    lcs_id = input('What is the LCS ID?'
                   '\n'
                   '--> ')

    for chart in chart_gatherer.gathered_charts:
        chart.set_active_worksheet(target=0)
        chart.active_worksheet.title = lcs_id

