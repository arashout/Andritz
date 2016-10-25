import os
import win32com.client as win32


def docC2txt(file_path):
    # Convert word .doc to .txt
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False

    doc = word.Documents.Open(file_path)
    txt_file_path = file_path.rstrip('doc') + 'txt'

    word.ActiveDocument.SaveAs(
        txt_file_path, FileFormat=win32.constants.wdFormatTextLineBreaks)
    word.Application.Quit(-1)

    return txt_file_path
