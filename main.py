import re
from datetime import datetime
from pathlib import Path

import pandas as pd

import codebook


def read_file(input_file: str):
    input_path = Path(input_file)
    if not input_path.exists():
        raise Exception('00', 'input path is not exists')
    if not input_path.is_file():
        raise Exception('01', 'input path is not a file')
    if not input_path.suffix == '.xlsx':
        raise Exception('02', 'input path extension not allow')
    return pd.read_excel(input_path, sheet_name=None)


def normalize(sheet_content):
    group_value = {}
    for group in range(6, 12):
        column = sheet_content.iloc[:, group]
        column_value = []
        for _, group_name in enumerate(codebook.group.keys()):
            if group_name in column.name:
                for key, value in enumerate(column):
                    group_items = {k: 0 for k, v in codebook.group[group_name].items()}
                    if not pd.isna(value):
                        rows = re.findall(r'[A-Z][0-9]+', value)
                        for _, item in enumerate(rows):
                            if item in group_items.keys():
                                group_items[item] = 1
                    row = list(sheet_content.iloc[key, 0:6]) + list(group_items.values())
                    column_value.append(row)
                df_column_value = pd.DataFrame(column_value)
                df_column_value.columns = list(sheet_content.columns[0:6]) + list(codebook.group[group_name])
                group_value[group_name] = df_column_value
                break
    return group_value


class Helper:

    def __init__(self) -> None:
        super().__init__()
        self.input_excel = None
        self.normalization = None
        self.result = None

    def export(self, output_file):
        output_path = Path(output_file)
        if output_path.is_dir():
            filename = f"{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            output_path = output_path.joinpath(filename)
        if output_path.exists():
            index = 1
            origin = output_path.stem
            while output_path.exists():
                output_path = output_path.with_name(f'{origin}_{index}{output_path.suffix}')
                index += 1
        output_path.parent.mkdir(parents=True, exist_ok=True)
        writer = pd.ExcelWriter(output_path)
        for key, value in self.normalization.items():
            value.to_excel(writer, key, index=False)
        writer.save()
        return output_path

    def run(self, input_file: str, output_file: str):
        self.input_excel = read_file(input_file)
        for sheet_name, sheet_content in self.input_excel.items():
            if '表單回應' not in sheet_name:
                continue
            self.normalization = normalize(sheet_content)
        output_path = self.export(output_file)
        return output_path


if __name__ == '__main__':
    helper = Helper()
    helper.run(f'./input/file1.xlsx', f'./output/result.xlsx')
