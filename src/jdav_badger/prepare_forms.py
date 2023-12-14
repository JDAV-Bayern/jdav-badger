import csv
import os
from openpyxl import load_workbook

# input csv file
INPUT_FILE = 'JLMarken Test.csv'

# template for output excel file
TEMPLATE_NAME = 'Jugendleitermarke Bestell Liste'
TEMPLATE_FILE = TEMPLATE_NAME + '.xlsx'

# column indexes in input csv file
GIVEN_NAME_COLUMN = 0
FAMILY_NAME_COLUMN = 1
CHAPTER_COLUMN = 2
BADGE_COLUMN = 3

# row to insert data at
INSERT_LINE = 5

# folder to collect output files in
OUTPUT_FOLDER = 'out'

# map badge status to full word
badgeMap = {
    'J': 'Ja',
    'B': 'Bedingt',
    'N': 'Nein',
    '': 'Nein',
}

result = {}

with open(INPUT_FILE, 'r', encoding='cp1252') as file:
    reader = csv.reader(file, delimiter=';')

    # skip header
    next(reader, None)

    for row in reader:
        stripped = [s.strip() for s in row]
        chapter = stripped[CHAPTER_COLUMN].replace('/', '_')
        badge = badgeMap[stripped[BADGE_COLUMN].upper()]

        # copy data
        data = [
            stripped[GIVEN_NAME_COLUMN],
            stripped[FAMILY_NAME_COLUMN],
            badge,
        ]

        # store data by chapter name
        if chapter in result:
            result[chapter].append(data)
        else:
            result[chapter] = [data]

# make sure output directory exists
if not os.path.isdir(OUTPUT_FOLDER):
    os.mkdir(OUTPUT_FOLDER)

for chapter in result.keys():
    # open worksheet
    workbook = load_workbook(TEMPLATE_FILE)
    worksheet = workbook.active

    # insert sufficient rows
    worksheet.insert_rows(INSERT_LINE, amount=len(result[chapter]))

    # write cells
    for row_index, row_data in enumerate(result[chapter], start=INSERT_LINE):
        for col_index, cell_value in enumerate(row_data, start=1):
            worksheet.cell(row=row_index, column=col_index).value = cell_value

    # save result to output folder
    workbook.save(OUTPUT_FOLDER + '/' + TEMPLATE_NAME + ' ' + chapter + '.xlsx')
