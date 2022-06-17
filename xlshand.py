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
            data = xlrd.open_workbook_xls(file_)
            table = data.sheets()[0]
            for colume_ in range(0,table.ncols,2):
                col_list = table.col_values(colume_, start_rowx=0, end_rowx=None)
                sum_list_key = sum_list_key + col_list
    all_index = list(set(sum_list_key))
    # for idx,key in enumerate(all_index):
    #     if len(key) == 0:
    #         all_index[idx] = "nan"
    all_index = list(filter(None,all_index))
    create_temp_db(all_index,conn)


def insert_data(user_dict,conn):
    '''插入数据
    :index_single:索引，单张表格的索引
    :data:单张表格的数据,顺序与index对应
    '''
    tablename ="test"
    #去掉空值
    values = list(filter(None,user_dict.values()))
    keys = list(filter(None,user_dict.keys()))

    question_marks = ','.join(list('?'*len(keys)))
    values = tuple(values)
    keys = ",".join(keys)
    conn.execute('INSERT INTO '+tablename+' ('+keys+') VALUES ('+question_marks+')', values)
    conn.commit()
    print(f"insert data success")

def read_xls_write_dict_2(file_list:list,conn):
    '''读取excel文件，将数据写入字典,然后通过函数写入表格'''
    for file in file_list:
            try:
                if file.endswith(".xls"):
                    print(f"正在提取{file}表格数据")
                    data = xlrd.open_workbook_xls(file)
                    table = data.sheets()[0]
                    sum_list_key = []
                    sum_list_value = []
                    for colume_ in range(0,table.ncols,2):
                        col_list = table.col_values(colume_, start_rowx=0, end_rowx=None)
                        col_list_value = table.col_values(colume_ + 1, start_rowx=0, end_rowx=None)
                        sum_list_key = sum_list_key + col_list
                        sum_list_value = sum_list_value + col_list_value 
                    mydict = dict(zip(sum_list_key,sum_list_value))
                    insert_data(mydict,conn)
            except:
                print(f"error when handle {file}")
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
        sheet.append(rows)
    workbook.save("output.xls")


if __name__ == "__main__":
    conn = sqlite3.connect(":memory:")
    file_list = os.listdir("./")
    iter_all_file(file_list,conn)
    read_xls_write_dict_2(file_list,conn)
    print("verify success")
    write2excel(conn)