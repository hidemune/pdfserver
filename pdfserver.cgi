#!/usr/bin/python3
'''
LibreOffice を Python で操作する

コマンドラインからのマクロの起動
soffice 'vnd.sun.star.script:args.py$show_args(1 , b)?language=Python&location=user'

pip install numpy
pip install pandas
pip3 install odfpy
pip3 install ezodf

libreoffice --headless --nologo --nofirststartwizard --convert-to pdf --outdir <output_path> <input_file.xlsx>  
'''

startRow = 4

import os,io,sys
import cgi
import subprocess
import ezodf

form = cgi.FieldStorage()

def getCellAddress(row , col):
    ret = ''
    dict = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    wk = col
    if wk <= 26:
        ret += dict[wk:wk+1]
        ret += str(row+startRow)
        return ret

class TestError(Exception):
    """XBRLの解析中にエラーが発生したことを知らせる例外クラス"""
    pass

try:
    import uuid
    u1 = str(uuid.uuid1())
    import datetime
    now = datetime.datetime.now()
    timename = now.strftime("%Y%m%d_%H%M%S%f") 
    fname = u1
    
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
    sys.stdin = io.TextIOWrapper(sys.stdin.buffer, encoding='utf-8')
    
    data = '''1行め文字列	数値	あああ
test	1234	あうあう
ほげ	3456	hogehoge
ほげ	3456	hogehoge
test	1234	あうあう
ほげ	3456	hogehoge
test	1234	あうあう
ほげ	3456	hogehoge
test	1234	あうあう
ほげ	3456	hogehoge
test	1234	あうあう
ほげ	3456	hogehoge
test	1234	あうあう
ほげ	3456	hogehoge
test	1234	あうあう
最後の行ほげ	3456	hogehoge
'''
    
    ods = ezodf.opendoc('calc_open.ods')
    
    sheet = ods.sheets[0]
    #sheet = ods.sheets['Sheet1']
    sheet.append_columns(3)
    
    lines = data.split('\n')
    sheet.append_rows(len(lines))
    for r,line in enumerate(lines):
        arr = line.split('\t')
        for c,itm in enumerate(arr):
            sheet[getCellAddress(r,c)].set_value(itm)
    
    #subprocess.run("echo 'test01'", shell=True)
    ods.saveas('output/'+fname+'.ods')
    #raise TestError('Test02')
    try:
        res = subprocess.run("sudo -u user libreoffice --headless --nologo --nofirststartwizard --convert-to pdf --outdir output output/" + fname + ".ods > /dev/null 2>&1", shell=True)
        #print ("Content-Type: text/html")
        #print ("")
        #print(str(res)) # CompletedProcess(args='sudo -u user libreoffice --headless --nologo --nofirststartwizard --convert-to pdf --outdir output output/eae3b1ce-6986-11ec-888f-008e257304c2.ods', returncode=0)
    except Exception as ex2:
        print ("Content-Type: text/html")
        print ("")
        print(str(ex)) 
    sys.stdout.buffer.flush()
    sys.stdin.buffer.flush()
    sys.stderr.buffer.flush()
    #raise TestError('Test03')
    with open(os.path.abspath('output/'+fname+'.pdf'), 'rb') as f:
        sys.stdout.buffer.write(b"Content-Type: application/pdf;\nContent-Disposition: inline; filename=Generated.pdf\n\n")
        sys.stdout.buffer.write(f.read())
except Exception as ex:
    sys.stdout.buffer.flush()
    print ("Content-Type: text/html")
    print ("")
    print(str(ex))
