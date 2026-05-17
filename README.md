# Office MCP Server

全球首批办公软件 MCP 工具集 — 同时支持 Excel + Word + PPT。

## 工具列表 (13个)

### Excel (5个)
- `excel_open` — 打开 Excel 文件
- `excel_read` — 读取指定区域数据
- `excel_write` — 写入单元格
- `excel_get_info` — 工作簿概况
- `excel_save` — 保存工作簿

### Word (4个)
- `word_open` — 打开 Word 文档
- `word_read_text` — 读取全部文本
- `word_insert_text` — 插入文字
- `word_save` — 保存文档

### PPT (4个)
- `ppt_open` — 打开演示文稿
- `ppt_list_slides` — 列出所有幻灯片
- `ppt_add_text_slide` — 添加文字页
- `ppt_save` — 保存

## 要求
- Windows + Microsoft Office (win32com)
- Python 3.10+

## 安装
```bash
pip install mcp pywin32
python server.py
```
