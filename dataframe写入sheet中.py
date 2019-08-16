# -*- coding: utf-8 -*-
import pandas as pd
#DataFrame
df1 = pd.DataFrame({'a': [1, 2, 3], 'b': [4, 5, 6]}, columns=['a', 'b'])
df2 = pd.DataFrame({'a': [7, 8, 9], 'b': [6, 5, 4]}, columns=['a', 'b'])
#指定Excel
writer = pd.ExcelWriter('aaa.xlsx')
#dataframe插入Excel
df1.to_excel(writer, sheet_name='sheet1', index=False)
df2.to_excel(writer, sheet_name='sheet2', index=False)
#保存
writer.save()
