import openpyxl
import numpy as np
import cv2
from PIL import Image


def separate_images(np_image, ocr_engine):

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
    address_total = ''
    total_total = ''
    date_total = ''
    company_total = ''
    total = ''
    date = ''
    company = ''
    address = ''

    for i in range(len(extracted_data)):
        if 'ADDRESS' in extracted_data[i][0]:
            if 'B' in extracted_data[i][0]:
                address = extracted_data[i][1]
            if 'I' in extracted_data[i][0]:
                address = address + ' ' + extracted_data[i][1]
        if 'DATE' in extracted_data[i][0]:
            if 'B' in extracted_data[i][0]:
                date = extracted_data[i][1]
            if 'I' in extracted_data[i][0]:
                date = date + ' ' + extracted_data[i][1]
        if 'COMPANY' in extracted_data[i][0]:
            if 'B' in extracted_data[i][0]:
                company = extracted_data[i][1]
            if 'I' in extracted_data[i][0]:
                company = company + ' ' + extracted_data[i][1]
        if 'TOTAL' in extracted_data[i][0]:
            if 'B' in extracted_data[i][0]:
                total = extracted_data[i][1]
            if 'I' in extracted_data[i][0]:
                total = total + ' ' + extracted_data[i][1]

        # choosing the longest output token
        if len(total) > len(total_total):
            total_total = total
        if len(date) > len(date_total):
            date_total = date
        if len(company) > len(company_total):
            company_total = company
        if len(address) > len(address_total):
            address_total = address

    # naming god
    address_total_date_company = []
    address_total_date_company.append(address_total)
    address_total_date_company.append(total_total)
    address_total_date_company.append(date_total)
    address_total_date_company.append(company_total)

    return address_total_date_company


def fill_exel_report(extracted_data, report_filename='exel_report.xlsx'):
    book = openpyxl.open(report_filename, read_only=False)
    sheet = book.worksheets[0]

    for extr_data in extracted_data:
        data = parse_extracted_data(extr_data)

        for i in range(100):
            if sheet['D' + str(68 + i)].value == None:
                sheet['D' + str(68 + i)] = data[2]
                sheet['H' + str(68 + i)] = data[3]
                sheet['L' + str(68 + i)] = data[1]
                break

    book.save(report_filename)
    book.close()
