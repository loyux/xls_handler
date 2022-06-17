import pandas as pd
data = pd.read_csv("testfile.txt", sep=" ")
data.to_excel("assas.xls", index=False)