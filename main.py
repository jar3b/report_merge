# Press the green button in the gutter to run the script.
import csv
import os
from pathlib import Path
from typing import List, Any

import xlrd


def opti(x: Any) -> Any:
    if not isinstance(x, str):
        return x
    return x.replace('\xa00', ' ')


def parse_files(in_path: str, out_filename: str) -> None:
    out_data = []
    headers_old_style = ['Артикул', 'Наименование товара', 'Ед. изм.', 'Кол-во', 'Цена с НДС, руб.',
                         'Сумма с НДС, руб', 'номер счета']

    headers_new_style = ['Артикул', 'Наименование товара', 'Ед. изм.', 'Кол-во', 'Цена с НДС, руб.',
                         'Цена без НДС, руб',
                         'Сумма без НДС, руб.', '% НДС', 'Сумма НДС, руб.', 'Сумма с НДС, руб', 'номер счета']

    # headers_new_style
    headers: List[str] | None = None

    for fn in Path(in_path).glob("*.xls*"):
        wb = xlrd.open_workbook(fn)
        sh = wb.sheet_by_index(0)
        bill_number = str(fn).split('.')[0]

        for rownum in range(sh.nrows):
            row = sh.row_values(rownum)
            try:
                int(row[0])
                data = [opti(x) for x in row if x]
                data.append(bill_number)
                out_data.append(data)

                if not headers:
                    for headers_style in [headers_old_style, headers_new_style]:
                        if len(headers_style) == len(data):
                            headers = headers_style

                    if headers is None:
                        print(f'Invalid headers size: {len(data)}')
                        return
            except IndexError:
                pass
            except ValueError:
                pass
            except TypeError:
                pass

        print(f'Got {len(out_data)} lines from {fn}')

    if headers is None:
        print('No data to process')
        return

    os.remove(Path(in_path, out_filename))
    with open(Path(in_path, out_filename), mode='w', newline='', encoding='cp1251') as csv_file:
        writer = csv.writer(csv_file, delimiter=';', quotechar='"', quoting=csv.QUOTE_ALL)
        writer.writerow(headers)
        writer.writerows(out_data)


if __name__ == '__main__':
    parse_files('.', 'scheta.csv')
