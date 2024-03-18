import openpyxl
import numpy as np
import cv2
from PIL import Image


def separate_images(np_image, ocr_engine):
    # TODO: To delete this func?

    bounds = ocr_engine.readtext(np_image)
    pil_image = Image.fromarray(np_image)

    dots_coner = []
    dots_sered = []

    for bound in bounds:  # сохранение двух точек бокса
        p0, p1, p2, p3 = bound[0]
        dots_coner.append(p0)
        dots_coner.append(p1)
        dot = [(p0[0] + p2[0]) / 2, (p0[1] + p2[1]) / 2]
        dots_sered.append(dot)

    k_ans = 0
    projection_dot = []

    for j in dots_coner:  # прорекция на прямую
        y = (k_ans * j[0] - k_ans * k_ans * j[1]) / (1 - k_ans * k_ans)
        if k_ans == 0:
            x = j[0]
        else:
            x = y / k_ans
        dot = [x, y]
        projection_dot.append(dot)

    x = []
    g = 0
    right = []
    left = []
    perep = [] + [0]

    for i in projection_dot:
        x = x + [i[0]]

    x.sort()

    l = x[len(x) - 1]
    r = l / len(x)

    for i in range(0, l, 15):  # подсчет количества точек возле прямой
        rcount = 0
        lcount = 0
        for j in x:
            if (j < i + 15) and (j >= i - 15):
                lcount = lcount + 1
        left = left + [lcount]

    for i in range(len(left)):  # нахождение самых больших скоплений точек
        if left[i] >= 17:
            perep = perep + [i]

    X = np.array(dots_coner)
    count = 0
    gr = []

    dotss = x

    x, y = pil_image.size
    perep = perep + [x / 15]
    for i in range(len(perep)):  # фильтрация
        if i != 0:
            if perep[i] - perep[i - 1] < 5:
                count = count + 1
            else:
                gr = gr + [(perep[i - 1] + perep[i - 1 - count]) / 2]
                count = 0
        if i == len(perep) - 1:
            gr = gr + [(perep[i] + perep[i - count]) / 2]

    corp_imgs = []
    for i in range(1, len(gr)):  # обрезка изображений
        corp_imgs = corp_imgs + [pil_image.crop((gr[i - 1] * 15, 0, gr[i] * 15, y))]

    # TODO: сделать фильтрацию мусора
    return corp_imgs, gr


def parse_extracted_data(extracted_data):
    # extracted_data is list of ( ()-TOKEN, WORD ) tuples

    # collect all sequences
    seqs_dict = {'address_seqs': [], 'company_seqs': [], 'date_seqs': [], 'total_seqs': []}

    curr_address = []
    curr_company = []
    curr_date = []
    curr_total = []

    # ugly
    for i in range(len(extracted_data)):
        # check for entity name
        if 'ADDRESS' in extracted_data[i][0]:
            # different cases for I-B-I-I, B-I-I, I-I-I, B-I-B-I-I
            if 'B' in extracted_data[i][0] and curr_address == []:
                curr_address = [extracted_data[i][1]]
            elif 'B' in extracted_data[i][0] and curr_address != []:
                seqs_dict['address_seqs'].append(curr_address)
                curr_address = [extracted_data[i][1]]
            elif 'I' in extracted_data[i][0] and curr_address == []:
                curr_address = [extracted_data[i][1]]
            elif 'I' in extracted_data[i][0] and curr_address != []:
                curr_address.append(extracted_data[i][1])

        elif 'COMPANY' in extracted_data[i][0]:
            if 'B' in extracted_data[i][0] and curr_company == []:
                curr_company = [extracted_data[i][1]]
            elif 'B' in extracted_data[i][0] and curr_company != []:
                seqs_dict['company_seqs'].append(curr_company)
                curr_company = [extracted_data[i][1]]
            elif 'I' in extracted_data[i][0] and curr_company == []:
                curr_company = [extracted_data[i][1]]
            elif 'I' in extracted_data[i][0] and curr_company != []:
                curr_company.append(extracted_data[i][1])

        elif 'DATE' in extracted_data[i][0]:
            if 'B' in extracted_data[i][0] and curr_date == []:
                curr_date = [extracted_data[i][1]]
            elif 'B' in extracted_data[i][0] and curr_date != []:
                seqs_dict['date_seqs'].append(curr_date)
                curr_date = [extracted_data[i][1]]
            elif 'I' in extracted_data[i][0] and curr_date == []:
                curr_date = [extracted_data[i][1]]
            elif 'I' in extracted_data[i][0] and curr_date != []:
                curr_date.append(extracted_data[i][1])

        elif 'TOTAL' in extracted_data[i][0]:
            if 'B' in extracted_data[i][0] and curr_total == []:
                curr_total = [extracted_data[i][1]]
            elif 'B' in extracted_data[i][0] and curr_total != []:
                seqs_dict['total_seqs'].append(curr_total)
                curr_total = [extracted_data[i][1]]
            elif 'I' in extracted_data[i][0] and curr_total == []:
                curr_total = [extracted_data[i][1]]
            elif 'I' in extracted_data[i][0] and curr_total != []:
                curr_total.append(extracted_data[i][1])


    # append last sequence
    seqs_dict['address_seqs'].append(curr_address)
    seqs_dict['company_seqs'].append(curr_company)
    seqs_dict['date_seqs'].append(curr_date)
    seqs_dict['total_seqs'].append(curr_total)

    # select longest sequence by characters for each entity
    entities_dict = {'address': max([' '.join(seqs) for seqs in seqs_dict['address_seqs']], key=len),
                     'company': max([' '.join(seqs) for seqs in seqs_dict['company_seqs']], key=len),
                     'date': max([' '.join(seqs) for seqs in seqs_dict['date_seqs']], key=len),
                     'total': max([' '.join(seqs) for seqs in seqs_dict['total_seqs']], key=len)}

    print(seqs_dict['total_seqs'])

    return entities_dict


def fill_exel_report(extracted_data_list, report_filename='exel_report.xlsx'):
    # extracted data list contains data for every receipt on the image
    book = openpyxl.open(report_filename, read_only=False)
    sheet = book.worksheets[0]

    # list for storing number of table rows of receipts
    table_numbers = []

    # 68 is a magic number
    for i, row in enumerate(sheet):
        if number := sheet['B' + str(68 + i)].value:
            table_numbers.append(number)
        else:
            break

    table_number_border = max(table_numbers)

    # row 82-86 is subscription structure (should copy it)


    for data in extracted_data_list:
        entities_dict = parse_extracted_data(data)

        # TODO: STOPPED HERE HELLLOOOOO
        filled_row = False
        for i in range(table_number_border):
            if sheet['D' + str(68 + i)].value == None:
                sheet['D' + str(68 + i)] = entities_dict['date']
                sheet['H' + str(68 + i)] = entities_dict['company']
                sheet['L' + str(68 + i)] = entities_dict['total']
                # TODO: в будущем научиться детектить и заполнять номер чека
                filled_row = True
                break
        if not filled_row:
            pass
            # TODO: should extend table and also lower subscription structure

    # book.save(report_filename)
    # book.close()


if __name__ == '__main__':

    nerv = [('I-ADDRESS', '<s>'),
            ('B-ADDRESS', '115114'),
            ('I-ADDRESS', 'Mockba'),
            ('I-ADDRESS', 'Площадь'),
            ('I-ADDRESS', 'Павелечкого'),
            ('I-ADDRESS', 'вокзала'),
            ('I-ADDRESS', ''),
            ('I-ADDRESS', ','),
            ('I-ADDRESS', 'la'),
            ('B-ADDRESS', '115114,'),
            ('I-ADDRESS', 'r.Mockba,'),
            ('I-ADDRESS', 'nл.Павеле'),
            ('I-ADDRESS', 'ECC'),
            ('I-ADDRESS', ''),
            ('I-ADDRESS', '4A,'),
            ('B-DATE', ':31.10.2023'),
            ('I-ADDRESS', 'Aомодейово'),
            ('I-ADDRESS', '189'),
            ('I-ADDRESS', 'Павелeчkий'),
            ('B-DATE', '31.10.23'),
            ('B-TOTAL', ''),
            ('B-TOTAL', '=500.00'),
            ('B-TOTAL', '=3000.00'),
            ('I-COMPANY', 'OOO'),
            ('I-COMPANY', 'Слава'),
            ('I-COMPANY', 'Инфик'),
            ('B-COMPANY', 'Мышь')]

    result = parse_extracted_data(nerv)
    for k, v in result.items():
        print(f'{k}:{v}')

    fill_exel_report([result])

