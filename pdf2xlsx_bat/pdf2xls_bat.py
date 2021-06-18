# -*- coding: utf-8 -*-

"""
Created on Wed Jun 17 2021
@author: Chester.Liu
"""
import os
from configparser import ConfigParser
from io import open
from concurrent.futures import ProcessPoolExecutor

import pdfplumber
import xlwt

def pdf_to_xlsx(pdf_file_path, xlsx_file_path):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Sheet1')
    i = 0

    pdf = pdfplumber.open(pdf_file_path)
    for page in pdf.pages:                   
        for table in page.extract_tables():
            for row in table:
                for j in range(len(row)):
                    sheet.write(i, j, row[j])
                i += 1

    pdf.close()

    workbook.save(xlsx_file_path)
    print('写入excel成功，保存位置：' , xlsx_file_path)

def main():
    config_parser = ConfigParser()
    config_parser.read('config_pdf2xlsx.cfg')
    config = config_parser['default']

    tasks = []
    with ProcessPoolExecutor(max_workers=int(config['max_worker'])) as executor:
        for file in os.listdir(config['pdf_folder']):
            extension_name = os.path.splitext(file)[1]
            if extension_name != '.pdf':
                continue
            file_name = os.path.splitext(file)[0]
            pdf_file = config['pdf_folder'] + '/' + file
            xlsx_file = config['xlsx_folder'] + '/' + file_name + '.xlsx'
            print('正在处理: ', file)
            result = executor.submit(pdf_to_xlsx, pdf_file, xlsx_file)
            tasks.append(result)
    while True:
        exit_flag = True
        for task in tasks:
            if not task.done():
                exit_flag = False
        if exit_flag:
            print('完成')
            exit(0)


if __name__ == '__main__':
    main()
