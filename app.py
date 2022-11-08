#!/usr/bin/env python3
from flask import Flask, request,  redirect, render_template,session
from werkzeug.utils import secure_filename
import os,json
import string
import utils
#coding=utf-8

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = '/home/barcodeupload/'
app.secret_key='QWERTYUIOP'#对用户信息加密
app.config['SESSION_COOKIE_HTTPONLY'] = False
app.config['SESSION_COOKIE_NAME'] = 'barcodeGen'
app.config['JSON_AS_ASCII'] = False

@app.route('/')
def rootdir():
    return redirect('/upload')
    #return 'hello world'

@app.route('/upload', methods=['GET','POST'])
def upload():
    if request.method == 'POST':
        file_dict = request.files.to_dict()
        print("file_dict:",file_dict)
        src = request.form.get('src',None)
        dest = request.form.get('dest',None)
        if file_dict == {} or src not in list(string.ascii_uppercase) or dest not in list(string.ascii_uppercase):
            #return {"error":"文件和信息源列以及条形码目标列不能为空".encode('utf-8').decode('utf-8')}
            print("文件和信息源列以及条形码目标列不能为空")
            return redirect('/upload')
        if src == dest:
            #return {"error":"文件和信息源列以及条形码目标列不能为空".encode('utf-8').decode('utf-8')}
            print("信息源列/条形码目标列不能相同")
            return redirect('/upload')
        excelfile = file_dict.get("excelfile")
        excelfile.save(excelfile.filename)

        print("src=%s dest=%s"%(src,dest))
        return utils.xlsx_process(excelfile.filename,src,dest)
    else:
        return render_template('upload.html')

if __name__ == "__main__":
    app.run(processes=True,debug=True, host="0.0.0.0",port="3213")
