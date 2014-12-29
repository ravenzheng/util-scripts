# -*- coding: UTF-8 -*-
"""
Convert word documents to utf-8 encoded txt files
"""

import os
import win32com.client


source_dir = u'E:\\已完成摘要\\'.encode('gbk')
target_dir = 'E:\\resource\\\word\\'


# scan directory to get all the word documents
def scan(src_dir):
    word_files = [];
    for root, dirs, files in os.walk(src_dir):
        for file_name in files:
            ext = os.path.splitext(file_name)[1]
            if ext == '.doc':
                word_files.append(os.path.join(root, file_name));
    return word_files;


# convert word document to utf-8 encoded txt file
def convert(files):
    word_app = win32com.client.gencache.EnsureDispatch('Word.Application')
    word_app.Visible = True
    for doc_file_path in files:
        print 'Converting %s' % doc_file_path
        file_name = os.path.basename(doc_file_path)
        doc = word_app.Documents.Open(doc_file_path)
        txt_file_path = os.path.join(target_dir, file_name[:-3] + 'txt')
        doc.SaveAs(txt_file_path,
                   FileFormat=2,    # WdSaveFormat, wdFormatText
                   Encoding=65001,  # MsoEncoding, msoEncodingUTF8
                   LineEnding=2)    # WdLineEndingType, wdLFOnly
        doc.Close()
    word_app.Quit()


if __name__ == '__main__':
    files = scan(source_dir)
    convert(files)
