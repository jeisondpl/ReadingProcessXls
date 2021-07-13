# from os import DirEntry
from openpyxl import Workbook
from functions import *
import time
from multiprocessing import Pool,cpu_count, Manager

directoryout = "C:\Prueba\Reading.xlsx"
dirJuan = "F:\CORALES BASE DE DATOS\juanCorales"
dirLocal  = "C:\Prueba"

pathsDir = explorar(dirLocal)

def listener(list):
    print("SAVING XD")
    book = load_workbook(directoryout)
    sheet = book['Sheet']
    sheet.append(list)
    book.save(directoryout)


if __name__ == '__main__':
    regList = []
    p = Pool(cpu_count() - 2)
    q = Manager().Queue()
    watcher = p.apply_async(listener, (q,))
    jobs = [p.apply_async(ReadAllRow, (path, q)) for path in pathsDir]

    for job in jobs:
        job.get()
    
    q.put('kill')
    p.close()
    p.join()

    # rows = 1  
    # for reg in regList:
    #     col = 2
    #     for r in reg:
    #         sheet.cell(row=rows, column=col).value = r
    #         col += 1
    #     rows += 1
