#这个程序为web端主程序
from flask import Flask, request, render_template, send_file # 导入 Flask 和 request 模块
import json
import pandas as pd
import os
import docx
def file_send(file_path):
    with open(file_path, 'rb') as f:
        while 1:
            data = f.read(20 * 1024 * 1024)  # per 20M
            if not data:
                break
            yield data
def loadlist():
   with open('file_list.txt', 'r') as f: #读取每行的文件名，这个文件里放的就是目前用户上传的文件，将每一行存储在字典结构里
       line = f.readlines()
       file_list = {}
       for i, line in enumerate(line):
           file_list[i+1] = line.strip()
       return file_list
excel_list=None
app = Flask(__name__) # 创建一个 Flask 应用
@app.route('/') # 定义根路由
def hello():
    #返回网页文件
   return render_template('index.html')
@app.route('/login',methods=['POST'])
def login():
    print(request.form)
    login_operation=request.form['login_operation']
    if login_operation=='登录':
        username = request.form['username']
        password = request.form['password']
        with open('users.json', 'r') as f:
            users = json.load(f)
        for user in users:
            if user["username"] == username and user["password"] == password:
                # 读取txt文件中的数据
                with open('file_list.txt', 'r') as f:
                    data = f.readlines()
                # 处理数据，生成下拉列表选项
                options = []
                for i, d in enumerate(data):
                    options.append({'value': i + 1, 'text': d.strip()})
                return render_template('edit.html',options=options)
        for user in users:
            if user["username"] != username or user["password"] != password:
                return '登录失败请重试'
    elif login_operation=='注册':
        return render_template('register.html')
    #登录到服务器，获取用户名和密码
@app.route('/register',methods=['POST'])
def register():
    name = request.form['name']
    phone = request.form['phone']
    username = request.form['username']
    password = request.form['password']
    with open('users.json', 'r') as f:
        users = json.load(f)
    user = {
        "name": name,
        "phone": phone,
        "username": username,
        "password": password
    }
    users.append(user)
    with open('users.json', 'w') as f:
        json.dump(users, f)
    return render_template('index.html')
@app.route('/edit',methods=['POST'])
def edit():
    global excel_list
    print(request.form)
    data_select=request.form['data_select']
    edit_operation=request.form['edit_operation']
    if edit_operation=='查看':
       data_select = int(data_select)  # 字典中存储的key索引为int类型，需要转换
       storege_name = loadlist()
       print(storege_name)
       #excel_header=1 #可以用header函数来跳过excel表格的第一行
       file_name = './uploads/' + storege_name[data_select]
       print(type(file_name))
       if file_name.endswith('.xlsx'):
           # 打开xlsx文件
           excel_list = pd.read_excel(file_name) # 读取excel文件 #这一行要主要随时有可能给data_select+1
           # 进行操作
           return render_template('check_excel.html', excel_list=excel_list.to_html())
       if  file_name.endswith('.docx'):
           # 打开docx文件
           excel_list = docx.Document(file_name)  #这里读取的是docx的word文档文件，但是依然保存在了excel_list这个变量中
           # 获取文本内容
           text = "\n".join([para.text for para in excel_list.paragraphs])
           return render_template("check_word.html", text=text)
           # 进行操作
       else:
           # 文件格式不支持
           return '文件格式不支持'
    elif edit_operation=='上传文件':
        return render_template('upload.html')
    elif edit_operation=='删除文件':
        with open('file_list.txt', 'r') as f:
            file_list = f.read().splitlines()
        data_select = int(data_select)  # 字典中存储的key索引为int类型，需要转换
        storge_name = loadlist() #这是个字典变量来存放名字
        # excel_header=1 #可以用header函数来跳过excel表格的第一行
        delete_1 = './uploads/' + storge_name[data_select]
        os.remove(delete_1)
        # 从文件名列表中删除该文件名
        file_list.remove(storge_name[data_select])
        # 更新txt文件
        with open('file_list.txt', 'w') as f:
            f.write('\n'.join(file_list))
        # 返回成功信息
        with open('file_list.txt', 'r') as f:
            data = f.readlines()
            # 处理数据，生成下拉列表选项
        options = []
        for i, d in enumerate(data):
            options.append({'value': i + 1, 'text': d.strip()})
        return render_template('edit.html', options=options)
    elif edit_operation == '下载文件':
        # 读取txt文件中的数据
        with open('file_list.txt', 'r') as f:
            data = f.readlines()
        # 处理数据，生成下拉列表选项
        options = []
        for i, d in enumerate(data):
            options.append({'value': i + 1, 'text': d.strip()})
        return render_template('download.html', options=options)
@app.route('/download', methods=['GET','POST'])
def download():
    dataselect = request.form['filename']
    storege_name = loadlist()
    data_select = int(dataselect)
    filename= storege_name[data_select]
    print(filename)
    filepath = './uploads/'
    if filename.endswith('.docx'):
        mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        return send_file(filepath + filename, mimetype=mimetype, as_attachment=True)
    elif filename.endswith('.doc'):
        mimetype = 'application/msword'
        return send_file(filepath + filename, mimetype=mimetype, as_attachment=True)
    elif filename.endswith('.xlsx'):
        mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return send_file(filepath + filename, mimetype=mimetype, as_attachment=True)
    else:
        return '该文件类型不支持下载，目前仅支持doc,docx,xlsx'
@app.route('/upload', methods=['POST'])
def upload():
    # 获取上传的文件
    file = request.files['file']
    filename = file.filename
    # 读取txt文件中的数据
    with open('file_list.txt', 'r') as f:
        data = f.readlines()
        # 检查是否已经存在相同的文件名
        if filename in data:
            return '文件已存在，请先删除原文件，然后再次上传'
        else:
            # 保存文件到本地
            file.save(os.path.join(os.getcwd(), 'uploads', filename))
            # 将文件名保存到一个文件中
            with open('file_list.txt', 'a') as f:
                f.write('\n'+filename)
            # 处理数据，生成下拉列表选项
            # 读取txt文件中的数据
            with open('file_list.txt', 'r') as f:
                data = f.readlines()
            options = []
            for i, d in enumerate(data):
                options.append({'value': i + 1, 'text': d.strip()})
            return render_template('edit.html', options=options)

@app.route('/search',methods=['POST'])
def search():
    global excel_list
    excel_operation=request.form['excel_operation']
    #excel_list = pd.read_excel('py人员信息标签模版x2.xlsx')  # 读取excel文件
    if excel_operation=='确定筛选':
       target = request.form['select']
       columns= request.form['columns']
       result=excel_list.loc[excel_list[columns]==target]
       #result = excel_list.loc[(excel_list[target_column1] == target_data1) & (excel_list[target_column2] == target_data2)]  # 查找符合两种条件的数据
       print(type(result))
       return render_template('check2_excel.html', result=result.to_html())
    if excel_operation == '修改指定数据':
        hang_num= request.form['edit']
        hang_num = int(hang_num) - 1
        # 获取表单数据
        input_data = {}
        for col in excel_list.columns:
            input_data[col] = request.form[col]

        # 处理表单数据
        for col in excel_list.columns:
            excel_list.loc[hang_num, col] = input_data[col]

        # 渲染表格到网页端
        return render_template('table.html', table=excel_list.to_html())
    rows = []
    for col in excel_list.columns:
        rows.append(f'<tr><td>{col}</td><td><input type="text" name="{col}"></td></tr>')
    form_html = f'<form method="post">{"".join(rows)}<input type="submit" value="提交"></form>'
    return form_html

if __name__ == '__main__':
    app.run()  # 运行 Flask 应用