# Python处理excel脚本

### 1.功能介绍

将如下格式的excel提取到一张表格的一行中，第一行取自所有表格的索引

i>运行目录：程序所在目录

ii>运行前安装:pip install -r requirments.txt

iii>运行命令 python xlshand.py

iv>输出文件为 output.xls

提取当前文件夹下形如的xls文件，将其按照索引和值输出到一个汇总表，汇总表每行代表一张表的数据

1列为索引，2列为值，依次类推

| **idx1** | **1** | **idx2** | **2** |
| -------------- | ----------- | -------------- | ----------- |
| **idx3** | **3** | **idx4** | **4** |
| **idx5** | **5** |                |             |
|                |             | **idx6** | **6** |

如此类结构，第一列为索引，第二列为值的格式，将其统一输出到一个excel中，形如：

| **idx1** | idx2 | **idx3** | idx4 | idx5 | idx6 |
| -------------- | ---- | -------------- | ---- | ---- | ---- |
| 1              | 2    | 3              | 4    | 5    | 6    |

多张表形如

| **idx1** | idx2        | idx3 | idx4 | idx5 | idx6 |
| -------------- | ----------- | ---- | ---- | ---- | ---- |
| **1**    | **2** | 3    | 4    | 5    | 6    |
|                | **k** |      | k    |      | k    |

## 2.测试

## 3.扩展

执定行或列为索引，与对应的值，进行提取

## PS:目前已知的问题

1.对于不存在索引存在值的excel格式报错(已解决)

| **idx1** | **1** | idx2 | 2 |
| -------------- | ----------- | ---- | - |
|                | **3** | idx4 | 4 |

2.对于存在索引不存在值的excel格式报错已解决)

| idx1 | 1 |
| ---- | - |
| idx3 |   |

3.如果索引为1，2，3等数字类型报错

str(int)无法解决

Error

| idx1 | 1 |
| ---- | - |
| 2    | 2 |

以下可以处理

| idx1  | value1 |
| ----- | ------ |
| "123" |        |
