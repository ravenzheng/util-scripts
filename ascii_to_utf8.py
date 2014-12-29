# -*- coding: UTF-8 -*-
"""
Convert gbk encoded txt files to utf-8 encoded txt files
"""

import os

source_dir = 'E:\\resource\\final2\\'
target_dir = 'E:\\resource\\final3\\'


# scan directory to get all the txt files
def scan(src_dir):
    txt_files = [];
    for root, dirs, files in os.walk(src_dir):
        for file_name in files:
            ext = os.path.splitext(file_name)[1]
            if ext == '.txt':
                txt_files.append(os.path.join(root, file_name));
    return txt_files;


# convert gbk encoded file to utf-8 encoded file
def convert(files):
    for file_path in files:
        f = open(file_path, 'r')
        file_name = os.path.basename(file_path)
        new_file_path = os.path.join(target_dir, file_name)
        content = f.read();
        f.close()
        new_content = content.decode('gbk').encode('UTF-8')
        print 'converting %s' % file_name
        new_f = file(new_file_path, 'w')
        new_f.write(new_content)
        new_f.close()


if __name__ == '__main__':
    files = scan(source_dir)
    convert(files)
