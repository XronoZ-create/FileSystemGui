import openpyxl
import csv


def get_kontragent():
    kontragents = {}

    wb = openpyxl.load_workbook('kontragent.xlsx')
    sheet = wb.active
    i = 0
    for row in sheet.rows:
        i += 1
        if i == 1:
            continue
        name = row[0].value

        files_nonredact = row[1].value.split(';')
        files = []
        for a in files_nonredact:
            if a[0] == ' ':
                files.append(a[1:])
            else:
                files.append(a)

        files_nonredact_b = row[2].value.split(';')
        files_b = []
        for a_b in files_nonredact_b:
            if a_b[0] == ' ':
                files_b.append(a_b[1:])
            else:
                files_b.append(a_b)

        kontragents[name] = [files, files_b]

    return kontragents

def get_documents():
    documents = []

    wb = openpyxl.load_workbook('files.xlsx')
    sheet = wb.active
    for row in sheet.rows:
        name = row[0].value
        folder = row[1].value
        obligation = row[2].value
        obligation_b = row[3].value

        documents.append( {'name': name, 'folder': folder, 'obligation': obligation, 'obligation_b': obligation_b} )

    return documents[1:]

def list_to_csv(list_export):
    with open('export_data.csv', mode='w') as csv_file:
        fieldnames = list_export[0].keys()
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames, delimiter=';', lineterminator='\n')

        writer.writeheader()
        for save_cell_dict in list_export:
            writer.writerow(save_cell_dict)