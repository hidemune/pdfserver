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

templatepath = '/home/user/pdfserver_template/'
datapath = '/home/user/BkUp/pdfsv_data/'

import os,io,sys
import cgi
import subprocess
import ezodf

dataD = '''1行め文字列	数値	あああ
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

form = cgi.FieldStorage()
data = form.getvalue('data',dataD)
startrow = form.getvalue('startrow','1')
template = form.getvalue('template','')

startRow = int(startrow)
template = template.replace('/','').replace('\\','')

html = '''Content-Type: text/html; charset=UTF-8

<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8">
</head>

<body>
<form action="pdfserver.cgi" method="post">
<textarea name="data">
%s
</textarea><br>
startrow:
<input type="text" name="startrow" value="4"><br>
template:
<input type="text" name="template" value="template_sample.ods"><br>
<input type="submit" value="submit">
</form>
<style>
textarea {
  min-width: 600px;
  min-height: 500px;
}
</style>
</body>
</html>
'''

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

import pathlib

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
    
    #if ( os.environ['REQUEST_METHOD'] == "GET" ):
    if form.getvalue('data','') == '':
        print( html % (data,) )
        sys.exit(0)
    lf = templatepath + '/' + template + '.lock'
    while True:
        lockf = pathlib.Path(lf)
        if not lockf.exists():
            break
    lockf = pathlib.Path(lf)
    lockf.touch()
    ods = ezodf.opendoc(templatepath + '/' + template)
    sheet = ods.sheets[0]
    #sheet = ods.sheets['Sheet1']
    #sheet.append_columns(3)
    lines = data.split('\n')
    for r,line in enumerate(lines):
        if line.strip() == "":
           break
        arr = line.split('\t')
        for c,itm in enumerate(arr):
            if r == 0:
                sheet.append_columns(1)
            sheet.append_rows(1)
            sheet[getCellAddress(r,c)].set_value(itm.strip())
    #sheet.reset(size=(50, 10))
    #subprocess.run("echo 'test01'", shell=True)
    ods.saveas(datapath + '/' + fname+'.ods')
    #raise TestError('Test02')
    try:
        res = subprocess.call("sudo -u user libreoffice --headless --nologo --nofirststartwizard --convert-to pdf --outdir " + datapath + " " + datapath + "/" + fname + ".ods > /dev/null 2>&1", shell=True)
    except Exception as ex2:
        print ("Content-Type: text/html")
        print ("")
        print(str(ex))
        lockf.unlink()
    sys.stdout.buffer.flush()
    sys.stdin.buffer.flush()
    sys.stderr.buffer.flush()
    #lockf.unlink()
    #raise TestError('Test03')
    with open(os.path.abspath(datapath + '/' + fname + '.pdf'), 'rb') as f:
        sys.stdout.buffer.write(b"Content-Type: application/pdf;\nContent-Disposition: inline; filename=Generated.pdf\n\n")
        sys.stdout.buffer.write(f.read())
except Exception as ex:
    sys.stdout.buffer.flush()
    print ("Content-Type: text/html")
    print ("")
    print(str(ex))
finally:
    if lockf.exists():
        lockf.unlink()
