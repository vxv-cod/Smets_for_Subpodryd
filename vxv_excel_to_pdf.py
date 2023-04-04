
import os
import pathlib
import win32com.client


def all_files(directory):
    for root, _, files in os.walk(directory):
        for file in files:
            yield os.path.join(root, file)

def GO(directory, sig, ui):
    app = win32com.client.Dispatch("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False

    files = len(os.listdir(path=directory))
    print(files)
    faillist = all_files(directory)


    for index_, fail in enumerate(faillist):
        xlsx = pathlib.Path(fail)
        if xlsx.suffix == ".xlsx":
    
            fn = fail.rsplit("\\", 1)[1]
            sig.signal_label.emit(ui.label, f'Публикация файлов в PDF {index_} из {files} : {fn}')

            xlsx_dir = xlsx.parent
            xlsx_dir = str(xlsx_dir)
            basename = xlsx.stem
            basename = str(basename)
            output_file = xlsx_dir + "/" + basename + ".pdf"
            book = app.Workbooks.Open(xlsx)
            xlTypePDF = 0
            book.ExportAsFixedFormat(xlTypePDF, output_file)

        nomerfail = index_ + 1
        countfail = files
        proc = round(nomerfail / countfail * 100)
        progressBar = ui.progressBar_1
        sig.signal_Probar.emit(progressBar, proc)
        
        '''Закрыть файл без сохранения'''
        book.Close(False)

    sig.signal_label.emit(ui.label, 'Готово. Создана папка " Result " рядом с иходной')

    app.Quit()



