import pandas as pd

sales_df = pd.read_excel('https://github.com/chris1610/pbpython/blob/master/data/sample-salesv3.xlsx?raw=true')
sales_summary = sales_df.groupby(['name'])['ext price'].agg(['sum', 'mean'])
# Reset the index for consistency when saving in Excel
sales_summary.reset_index(inplace=True)
writer = pd.ExcelWriter('sales_summary.xlsx', engine='xlsxwriter')
sales_summary.to_excel(writer, 'summary', index=False)
workbook = writer.book
workbook.filename = 'sales_summary.xlsm'
workbook.add_vba_project('vbaProject.bin')
writer.close()
