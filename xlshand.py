import xlrd
import os
import openpyxl
import sqlite3
import sys

def create_temp_db(index,conn):
    '''根据所有的索引去重后进行创建'''
    '''索引为数字会报错'''
    tablename ="test"
    keys = ",".join(index)
    #keys不能为int,str(int)无效
    print(f"创建新表格的索引为\n{keys}")

    conn.execute(f"CREATE TABLE IF NOT EXISTS {tablename} ({keys})")
    print("create table success")

def iter_all_file(file_list:list,conn):
    '''遍历所有文件，取得所有索引,对所有索引去重后，成为表字段'''
    '''增加idx与value不同时存在时候的处理逻辑'''
    all_index = []
    dict_lists = []
    index_list_keys_summary = []
    handled_file_list = []
    for file_ in file_list:
        if file_.endswith(".xls") or file_.endswith(".xlsx"):
            sum_list_key = []
            sum_list_value = []
            data = xlrd.open_workbook(file_)
            table = data.sheets()[0]
            if table.ncols % 2 == 1:
                print(f"excel文件{file_}的列为单数，索引与值不匹配，请调整文件结构")
                sys.exit(1)
                #有一个单数行为table.ncols
            for colume_ in range(0,table.ncols,2):
                col_list = table.col_values(colume_, start_rowx=0, end_rowx=None)
                col_list_value = table.col_values(colume_ + 1, start_rowx=0, end_rowx=None)
                sum_list_key = sum_list_key + col_list
                sum_list_value = sum_list_value + col_list_value
                index_list_keys_summary = col_list + index_list_keys_summary
            mydict_for_index = dict(zip(sum_list_key,sum_list_value))
            dict_lists.append(mydict_for_index)
            handled_file_list.append(file_)
    all_index = list(set(index_list_keys_summary))
    all_indexes = list(filter(None,all_index))
    create_temp_db(all_indexes,conn)
    return dict_lists,handled_file_list

def insert_data(user_dict,conn):
    '''插入数据
    :index_single:索引，单张表格的索引
    :data:单张表格的数据,顺序与index对应
    '''
    tablename ="test"
    #去掉字典的空键
    iter_list_keys = list(user_dict.keys())
    for ele in iter_list_keys:
        if ele == "":
            user_dict.pop(ele)
    # values = list(filter(None,user_dict.values()))
    # keys = list(filter(None,user_dict.keys()))
    # str_keys = list(map(lambda x:str(x), list(user_dict.keys())))
    # str_values = list(map(lambda x:str(x), list(user_dict.values())))
    question_marks = ','.join(list('?'*len(user_dict.keys())))
    values = tuple(user_dict.values())
    keys = ",".join(user_dict.keys())
    conn.execute('INSERT INTO '+tablename+' ('+keys+') VALUES ('+question_marks+')', values)
    conn.commit()
    print("insert data success")

def dict_insert2sqlite(dict_list:list,conn,handed_file_list:list):
    '''将字典列表写入sqlite数据库'''
    for idx,dict_ in enumerate(dict_list):
        print(f"开始写入{handled_file_list[idx]}文件的数据")
        insert_data(dict_,conn)


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
    dict_list,handled_file_list = iter_all_file(file_list,conn)
    dict_insert2sqlite(dict_list,conn,handled_file_list)
    write2excel(conn)
    print("执行成功，文件保存为当前目录下output.xls")