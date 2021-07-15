import time

import httpx
import openpyxl


def get_oked_by_bin(bin):
    url = f'https://stat.gov.kz/api/juridicalusr/counter/gov/?bin={bin}&lang=ru'
    headers = {
        'Accept': 'application/json',
        'Referer': 'https://stat.gov.kz/jur-search/bin',
    }

    while True:
        response = httpx.get(url, headers=headers)

        if response.status_code == httpx.codes.OK:
            response_json = response.json()
            return response_json['obj']['okedCode']
        elif response.status_code == httpx.codes.BAD_REQUEST:
            break

        time.sleep(0.2)


def get_bin_column_index(worksheet):
    for index, cell in enumerate(next(worksheet.rows)):
        if cell.value == 'bin':
            return index


def main():
    workbook = openpyxl.load_workbook('companies.xlsx')
    worksheet = workbook.active
    bin_column_index = get_bin_column_index(worksheet)
    oked_column = worksheet.max_column + 1

    # inserting oked column header
    worksheet.cell(row=1, column=oked_column).value = 'oked'

    for row, row_cells in enumerate(worksheet.rows, start=1):
        # skipping header row
        if row == 1:
            continue

        bin = row_cells[bin_column_index].value
        oked = get_oked_by_bin(bin)
        worksheet.cell(row=row, column=oked_column).value = oked or '-'

    workbook.save('companies_with_oked.xlsx')


if __name__ == '__main__':
    main()
