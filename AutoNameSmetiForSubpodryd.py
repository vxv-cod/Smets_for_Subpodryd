'''Определяем тип сметы, дописываем его в название файла и размещаем в 
структуру папок для загрузки через модуль Субподряд'''

import os
import pathlib
import shutil
import sys
from PyQt5 import QtCore, QtWidgets
from openpyxl import load_workbook
import openpyxl.utils
from collections import Counter


# from rich import print

import vxv_translitt_text
from Options import *
import vxv_excel_to_pdf
from okno_ui import Ui_Form

app = QtWidgets.QApplication(sys.argv)
Form = QtWidgets.QWidget()
ui = Ui_Form()
ui.setupUi(Form)
Form.show()

_translate = QtCore.QCoreApplication.translate
Title = 'Доработка структуры папок для Смет'
Form.setWindowTitle(_translate("Form", Title))

ui.tableWidget.setColumnWidth(0, 100)
ui.tableWidget.setColumnWidth(2, 170)


sig = Signals()


'''Добавляем элемент в таблицу и устанавливаем маску ввода'''
ipLineEdit = QtWidgets.QLineEdit(Form)
ipLineEdit.setFrame(False)
ipLineEdit.setInputMask('000.000.000;_')
ipLineEdit.setAlignment(QtCore.Qt.AlignCenter)
ui.tableWidget.setCellWidget(1, 2, ipLineEdit)



def GO(directory):
    progressBar = ui.progressBar_1
    sig.signal_Probar.emit(progressBar, 5)
    result_folder = directory.rsplit("\\", 1)[0] + "\\Result"
    try:
        os.mkdir(result_folder)
    except:
        try:
            shutil.rmtree(result_folder)
        except FileNotFoundError:
            sig.signal_label.emit(ui.label, '')
            return sig.signal_err.emit(Form, "Адресс не найден ! ! !")
        os.mkdir(result_folder)
    sig.signal_Probar.emit(progressBar, 10)

    '''
    Собираем: полный путь исходного файла
    '''
    fails_patch = []
    for Patch, dirs, files in os.walk(directory):
        if files != []:
            for name in files:
                xlsx = pathlib.Path(os.path.join(Patch, name))
                if xlsx.suffix == ".xlsx":
                    fails_patch.append(str(xlsx))
                    # print(xlsx)
        # print(files)
    # print(fails_patch)
    # lenfails_patch = len(fails_patch)

    # TipFails = {
    #     'ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ' : 'ЛР',
    #     'ОБЪЕКТНЫЙ СМЕТНЫЙ РАСЧЕТ' : 'ОС',
    #     'СВОДНЫЙ СМЕТНЫЙ РАСЧЕТ' : 'ССРСС',
    #     'ЛОКАЛЬНАЯ РЕСУРСНАЯ ВЕДОМОСТЬ' : 'ЛРВ',
    #     'ЛОКАЛЬНАЯ СМЕТА' : 'ЛР',
    #     'ОБЪЕКТНАЯ СМЕТА' : 'ОС'
    # }
    
    TipFails = {
        'ЛОКАЛЬН' : 'ЛР',
        'ОБЪЕКТН' : 'ОС',
        'СВОДН' : 'ССРСС',
        'РЕСУРСН' : 'ЛРВ',
    }


    dataTab = []
    for stolbec in range(4):
        if stolbec != 2:
            x = ui.tableWidget.item(1, stolbec).text()
        else:
            x = ipLineEdit.text()
        if x != '':
            dataTab.append(x)
        else:
            return sig.signal_err.emit(Form, f"Не заполнена {stolbec + 1}-ая ячейка таблицы")

    ShifrKO = "-".join(dataTab)
    # print(f'ShifrKO = {ShifrKO}')    

    '''Собираем структуру вложенных папок'''
    SH0 = dataTab[0]
    SH1 = dataTab[1]
    SH2 = [dataTab[2][:-4], dataTab[2][-3:]]
    SH3 = dataTab[3]
    SH = [SH0] + [SH1] + SH2 + [SH3]
    StructureFolder = "\\".join(SH)

    # Создаем структуру 
    endfolser = f'{result_folder}\\{StructureFolder}'
    os.makedirs(endfolser)

    newFailList = []
    tempname = []
    errorListfailRev = []
    nom = 0
    shifrinList = []

    for inde, filename in enumerate(fails_patch):

        fn = filename.rsplit("\\", 1)[1]
        sig.signal_label.emit(ui.label, 'Обработка файла: ' + fn)

        rev = 'None'
        shifrin = 'None'
        # nom = 1
        newtip = None
        wb = load_workbook(filename = filename)
        ws = wb.active

        '''-------------------------------------------------------------'''
        
        '''# Установка масштаба страницы для печати'''
        if ws.sheet_properties.pageSetUpPr is None:
            ws.sheet_properties.pageSetUpPr = openpyxl.SheetProperties()
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1000
        
        '''-------------------------------------------------------------'''

        rowslist = tuple(ws.values)[:35]
        CellsListPoisk = [str(cell) for row in rowslist for cell in row if cell != None]

        for cellvalue in CellsListPoisk:
            for tip in TipFails:
                if tip in str(cellvalue):
                    newtip = TipFails[tip]
                    # №
                    '''Поиск ревизии в документе'''
                    tmp_ = cellvalue.rsplit("-", 1)
                    shifrin = tmp_[0].split("№")[1]
                    shifrin = shifrin.strip(' ')
                    rev = tmp_[1]
                    rev = rev.strip('')
                    rev = vxv_translitt_text.GO(rev)
                    rev = rev.replace("RS", "rC")
                    rev = rev.replace("rS", "rC")
                    rev = rev.replace("rС", "rC")
                    if rev == '':
                        rev = 'None'
                        errorListfailRev.append(filename)

        tempname.append(f'{newtip}')
        shifrinList.append(f'{shifrin}')

        
        '''Считаем однинаковые значения в списке.
        Создается словарь, где ключ - это элемент, 
        значение - это количество элементов'''
        counter = Counter(tempname)

        '''# Присваиваем количество по порядку'''
        if newtip in counter:
            nom = counter[newtip]
        else:
            nom = 1

        # print(f'{inde} : {counter} - {newtip} - {nom}')

        nom_str = str(nom).rjust(3, '0')
        # NewNames = f'{ShifrKO}-{newtip}-{nom_str}-{rev} {shifrin}'
        NewNames = f"{ShifrKO}-{newtip}-{nom_str}-{rev.strip('')} {shifrin.strip('')}"
        newFailList.append(NewNames)

        rashireniefaila = '.' + filename.rsplit(".", 1)[1]
        fullfailNewName = endfolser + "\\" + NewNames +  rashireniefaila

        # shutil.copy2(filename, fullfailNewName)

        wb.save(fullfailNewName)

        nomerfail = inde + 1
        countfail = len(fails_patch)
        proc = round(nomerfail / countfail * 100)
        sig.signal_Probar.emit(progressBar, proc)

        wb.close()


    '''Вывод ошибки уведомления при не заполненных ревизиях в водкументе'''
    if errorListfailRev != []:
        text = ''''''
        for i in errorListfailRev:
            text += i + '\n'
        sig.signal_err.emit(Form, f'''Номер ревизии не найден в файле:\n{text}''')

    sig.signal_label.emit(ui.label, 'Подготовка к публикации файлов в PDF . . .')
    
    sig.signal_Probar.emit(progressBar, 0)

    vxv_excel_to_pdf.GO(endfolser, sig, ui)


'''----------------------------------------------------------------------------'''

NameProgram = "AutoNameSmetiForSubpodryd"

@thread
@startFun(NameProgram, Form, sig, [ui.pushButton], ui.progressBar_1, ui.label)
def start():
    directory = ui.plainTextEdit.toPlainText()
    if directory == '':
        sig.signal_label.emit(ui.label, '')
        return sig.signal_err.emit(Form, "Не указана исходная папка ! ! !")
    GO(directory)



# def loadListPoisk(fail):
#     '''Открываем файл с кодировкой для чтения'''
#     f = open(fail, 'r', encoding='utf-8')
#     '''Читаем весь файл целиком как текст'''
#     # text = f.read()
#     '''Читаем файл и разбивает строку на подстроки в зависимости от разделителя'''
#     text = f.read().split("\n") 
#     return text


# def openElements():
#     '''открытие файла как при двойном клике'''
#     os.startfile('Elements.ini')

# ipLineEdit.setText('666.777.888')

ui.plainTextEdit.clear()
ui.plainTextEdit.textChanged.connect(lambda : ChangedPT(ui.plainTextEdit))
ui.pushButton.clicked.connect(start)
# ui.pushButton_2.clicked.connect(openElements)

if __name__ == "__main__":
    # start()
    sys.exit(app.exec_())


