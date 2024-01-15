import pandas as pd


class ExcelParser:
    def __init__(self, file_path, columns):
        self.file_path = file_path
        # provide dict of excel columns
        self.columns = columns
        self.data = pd.read_excel(file_path)

    def parse_data(self):
        for index, row in self.data.iterrows():
            if isinstance(row, pd.Series):
                data = {key: row[self.columns[key]] for key in self.columns}
                yield data
            else:
                yield {}

