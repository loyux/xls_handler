import xlrd
import os
import xlwt
import openpyxl
import sqlite3
#处理的每张表格式一样，推广到不同长度索引的excel
#首先遍历所有文件，取得所有索引,返回索引的列表的列表
import platform

def create_temp_db(index,conn):
    '''根据所有的索引去重后进行创建'''
    tablename ="test"
    keys = ",".join(index)
    conn.execute("CREATE TABLE IF NOT EXISTS test (%s)" %(keys))
    conn.commit()
    print("create table success")


def iter_all_file(file_list:list,conn)->list:
    '''遍历所有文件，取得所有索引,对所有索引去重后，成为表字段'''
    all_index = []
    sum_list_key = []
    for file_ in file_list:
        if file_.endswith(".xls"):
            print(file_)
            data = xlrd.open_workbook_xls(file_)
            table = data.sheets()[0]
            for colume_ in range(0,table.ncols,2):
                col_list = table.col_values(colume_, start_rowx=0, end_rowx=None)
                sum_list_key = sum_list_key + col_list
    print(sum_list_key)
    all_index = list(set(sum_list_key))
    kk = list(filter(None, all_index))
    create_temp_db(kk,conn)


def insert_data(user_dict,conn,file_name):
    '''插入数据
    :index_single:索引，单张表格的索引
    :data:单张表格的数据,顺序与index对应
    '''
    tablename ="test"
    question_marks = ','.join(list('?'*len(user_dict)))
    values = tuple(user_dict.values())
    keys = ",".join(user_dict.keys())
    conn.execute('INSERT INTO '+tablename+' ('+keys+') VALUES ('+question_marks+')', values)
    conn.commit()
    print(f"insert {file_name} data success")

def read_xls_write_dict_2(file_list:list,conn):
    '''读取excel文件，将数据写入字典,然后通过函数写入表格'''
    for file in file_list:
            try:
                if file.endswith(".xls"):
                    data = xlrd.open_workbook_xls(file)
                    table = data.sheets()[0]
                    sum_list_key = []
                    sum_list_value = []
                    for colume_ in range(0,table.ncols,2):
                        col_list = table.col_values(colume_, start_rowx=0, end_rowx=None)
                        col_list_value = table.col_values(colume_ + 1, start_rowx=0, end_rowx=None)
                        sum_list_key = sum_list_key + col_list
                        sum_list_value = sum_list_value + col_list_value 
                    # print((sum_list_key,sum_list_value))
                    mydict = dict(zip(sum_list_key,sum_list_value))
                    insert_data(mydict,conn,file)
            except:
                continue

def write2excel(conn):
    '''将字典写入excel文件'''
    cur = conn.cursor()
    cur.execute("SELECT * from test")
    col_name = [tuple[0] for tuple in cur.description]
    title = col_name
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(title)
    for rows in cur.fetchall():
        # print(rows)
        sheet.append(rows)
    workbook.save("output.xls")



def transform(parent_path,out_path):
    '''将xlsx转换为xls'''
    fileList = os.listdir(parent_path)  #文件夹下面所有的文件
    num = len(fileList)
    for i in range(num):
        file_Name = os.path.splitext(fileList[i])   #文件和格式分开
        if file_Name[1] == '.xlsx':
            transfile1 = parent_path+'\\'+fileList[i]  #要转换的excel
            transfile2 = out_path+'\\'+file_Name[0]    #转换出来excel
            excel=win32.gencache.EnsureDispatch('excel.application')
            pro=excel.Workbooks.Open(transfile1)   #打开要转换的excel
            pro.SaveAs(transfile2+".xls", FileFormat=56)  #另存为xls格式
            pro.Close()
            excel.Application.Quit()









if __name__ == "__main__":
    conn = sqlite3.connect(":memory:")
    file_list = os.listdir("./")
    iter_all_file(file_list,conn)
    read_xls_write_dict_2(file_list,conn)
    print("verify success")
    write2excel(conn)