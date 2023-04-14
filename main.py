import os
import markdown
import win32com.client as win32
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import StringIO
import pandas as pd
import win32clipboard

# 要插入的Markdown文本
markdown_text = '''
| Header 1 | Header 2 | Header 3 |
| -------- | -------- | -------- |
| Cell 1   | Cell 2   | Cell 3   |
| Cell 4   | Cell 5   | Cell 6   |
'''
win32clipboard.OpenClipboard()
markdown_text = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT).decode('gbk')
win32clipboard.CloseClipboard()
# 使用markdown库将Markdown文本转换为HTML
html = markdown.markdown(markdown_text, extensions=['tables'])

# 使用pandas从HTML中读取表格数据
buffer = StringIO(html)
table_data = pd.read_html(buffer)[0]

# 创建WPS对象
wps = win32.gencache.EnsureDispatch('KWPS.Application')

# 获取当前文档对象
doc = wps.ActiveDocument

# 获取当前光标位置
selection = wps.Selection

# 将光标移到当前段落的结尾
selection.Collapse(win32.constants.wdCollapseEnd)

# 将光标向下移动一个段落
selection.Move(win32.constants.wdParagraph, 1)

# 将表格数据插入到Word文档中
table = doc.Tables.Add(selection.Range, table_data.shape[0]+1, table_data.shape[1])
for i, column in enumerate(table_data.columns):
    table.Cell(1, i+1).Range.Text = str(column)

for i, row in table_data.iterrows():
    for j, value in enumerate(row):
        table.Cell(i+2, j+1).Range.Text = str(value)

# 设置表格样式
table.Borders.Enable = True
table.AllowAutoFit = True
table.Rows.HeightRule = 1
table.Rows.Alignment = WD_ALIGN_PARAGRAPH.CENTER
for cell in table.Range.Cells:
    cell.VerticalAlignment = 1
    cell.Range.Paragraphs.Alignment = WD_ALIGN_PARAGRAPH.CENTER

last_row = table.Rows.Last
last_row.Select()
wps.Selection.MoveDown(Unit=win32.constants.wdLine, Count=1)

# 插入一个空行作为段落分隔符
doc.Range(wps.Selection.Range.End, wps.Selection.Range.End).InsertAfter('\n')

# 显示Word文档
wps.Visible = True
