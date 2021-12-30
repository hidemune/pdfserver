#!/usr/bin/python
# -*- coding: utf-8 -*-
'''
ファイルをアップロードする
'''

import cgi
import os, sys

html = '''Content-Type: text/html

<html>
<head>
  <meta http-equiv="Content-Type" content="text/html" charset="UTF-8" />
  <title>ODSファイルアップロード</title>
</head>
<body>
<h1>ODSファイルをアップロードする</h1>
<p>%s</p>
<form action="templateupload.cgi" method="post" enctype="multipart/form-data">
  <input type="file" name="file" accept=".ods">
  <input type="submit" />
</form>
</body>
</html>
'''

try:
    import msvcrt
    msvcrt.setmode(0, os.O_BINARY)
    msvcrt.setmode(1, os.O_BINARY)
except ImportError:
    pass

result = ''
form = cgi.FieldStorage()
if form.has_key('file'):
    item = form['file']
    if item.file:
        fout = file(os.path.join('/home/user/pdfserver_template/', item.filename), 'wb')
        while True:
            chunk = item.file.read(1000000)
            if not chunk:
                break
            fout.write(chunk)
        fout.close()
        result = 'アップロードしました。'

print html % result