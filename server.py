
"""Office MCP Server — Excel + Word + PPT 统一操控
全球首批办公软件 MCP 工具集
"""
from mcp.server.fastmcp import FastMCP
import pythoncom, os

mcp = FastMCP("Office MCP Server")

def _init_com():
    pythoncom.CoInitialize()

# ═══════════════════════════════════════════════
#  EXCEL 工具 (5个)
# ═══════════════════════════════════════════════

@mcp.tool()
def excel_open(filepath: str) -> str:
    """打开 Excel 文件，返回工作表信息
    
    Args:
        filepath: .xlsx/.xls 文件路径
    """
    _init_com()
    import win32com.client
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(filepath)
        sheets = [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)]
        return f"已打开：{os.path.basename(filepath)}\n工作表：{sheets}"
    except Exception as e:
        return f"Excel 打开失败：{str(e)}"

@mcp.tool()
def excel_read(sheet_name: str = "Sheet1", start_cell: str = "A1", end_cell: str = "Z100") -> str:
    """读取 Excel 指定区域的数据
    
    Args:
        sheet_name: 工作表名
        start_cell: 起始单元格，如 A1
        end_cell: 结束单元格，如 D20
    """
    _init_com()
    import win32com.client
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        wb = excel.ActiveWorkbook
        if wb is None:
            return "无活动工作簿"
        ws = wb.Sheets(sheet_name)
        rng = ws.Range(f"{start_cell}:{end_cell}")
        data = rng.Value
        if data is None:
            return f"{start_cell}:{end_cell} 无数据"
        # 格式化为表格
        if not isinstance(data, tuple):
            data = [[data]]
        lines = []
        for row in data[:50]:
            if row is None:
                continue
            if isinstance(row, (int, float, str)):
                row = [row]
            lines.append("\t".join(str(c) if c is not None else "" for c in row))
        summary = f"读取 {len(lines)} 行" + (f"（仅显示前50行，总计{len(data)}行）" if len(data) > 50 else "")
        return summary + "\n" + "\n".join(lines)
    except Exception as e:
        return f"读取失败：{str(e)}"

@mcp.tool()
def excel_write(sheet_name: str, cell: str, value: str) -> str:
    """写入数据到 Excel 单元格
    
    Args:
        sheet_name: 工作表名
        cell: 单元格位置，如 B3
        value: 要写入的值
    """
    _init_com()
    import win32com.client
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        wb = excel.ActiveWorkbook
        if wb is None:
            return "无活动工作簿"
        ws = wb.Sheets(sheet_name)
        ws.Range(cell).Value = value
        return f"已写入 {sheet_name}!{cell} = {value}"
    except Exception as e:
        return f"写入失败：{str(e)}"

@mcp.tool()
def excel_get_info() -> str:
    """获取当前 Excel 工作簿的概况"""
    _init_com()
    import win32com.client
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        wb = excel.ActiveWorkbook
        if wb is None:
            return "无活动工作簿"
        sheets_info = []
        for i in range(1, wb.Sheets.Count + 1):
            ws = wb.Sheets(i)
            used = ws.UsedRange
            rows = used.Rows.Count if used else 0
            cols = used.Columns.Count if used else 0
            sheets_info.append(f"  {ws.Name}: {rows}行 x {cols}列")
        return f"工作簿：{wb.Name}\n" + "\n".join(sheets_info) if sheets_info else "无工作表"
    except Exception as e:
        return f"获取失败：{str(e)}"

@mcp.tool()
def excel_save(filepath: str = "") -> str:
    """保存当前 Excel 工作簿"""
    _init_com()
    import win32com.client
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        wb = excel.ActiveWorkbook
        if wb is None:
            return "无活动工作簿"
        if filepath:
            wb.SaveAs(filepath)
        else:
            wb.Save()
        return f"已保存：{filepath or wb.FullName}"
    except Exception as e:
        return f"保存失败：{str(e)}"


# ═══════════════════════════════════════════════
#  WORD 工具 (4个)
# ═══════════════════════════════════════════════

@mcp.tool()
def word_open(filepath: str) -> str:
    """打开 Word 文档
    
    Args:
        filepath: .docx 文件路径
    """
    _init_com()
    import win32com.client
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(filepath)
        paragraphs = doc.Paragraphs.Count
        return f"已打开：{os.path.basename(filepath)}\n段落数：{paragraphs}"
    except Exception as e:
        return f"Word 打开失败：{str(e)}"

@mcp.tool()
def word_read_text() -> str:
    """读取当前 Word 文档的全部文本"""
    _init_com()
    import win32com.client
    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.ActiveDocument
        if doc is None:
            return "无活动文档"
        text = doc.Content.Text
        return f"文档：{doc.Name}\n字数：{len(text)}\n---\n{text[:2000]}" + ("..." if len(text) > 2000 else "")
    except Exception as e:
        return f"读取失败：{str(e)}"

@mcp.tool()
def word_insert_text(text: str, position: str = "end") -> str:
    """在 Word 文档中插入文字
    
    Args:
        text: 要插入的文字
        position: "end"(末尾) 或 "start"(开头)
    """
    _init_com()
    import win32com.client
    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.ActiveDocument
        if doc is None:
            return "无活动文档"
        rng = doc.Range(0, 0) if position == "start" else doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        rng.Text = text + ("\n" if position == "end" else "")
        return f"已在{position}位置插入 {len(text)} 字符"
    except Exception as e:
        return f"插入失败：{str(e)}"

@mcp.tool()
def word_save(filepath: str = "") -> str:
    """保存当前 Word 文档"""
    _init_com()
    import win32com.client
    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.ActiveDocument
        if doc is None:
            return "无活动文档"
        if filepath:
            doc.SaveAs(filepath)
        else:
            doc.Save()
        return f"已保存：{filepath or doc.FullName}"
    except Exception as e:
        return f"保存失败：{str(e)}"


# ═══════════════════════════════════════════════
#  PPT 工具 (4个)
# ═══════════════════════════════════════════════

@mcp.tool()
def ppt_open(filepath: str) -> str:
    """打开 PowerPoint 演示文稿
    
    Args:
        filepath: .pptx 文件路径
    """
    _init_com()
    import win32com.client
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        ppt.Visible = False
        pres = ppt.Presentations.Open(filepath)
        return f"已打开：{os.path.basename(filepath)}\n幻灯片数：{pres.Slides.Count}"
    except Exception as e:
        return f"PPT 打开失败：{str(e)}"

@mcp.tool()
def ppt_list_slides() -> str:
    """列出所有幻灯片及其内容概要"""
    _init_com()
    import win32com.client
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        pres = ppt.ActivePresentation
        if pres is None:
            return "无活动演示文稿"
        slides_info = []
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(i)
            shapes_count = slide.Shapes.Count
            # 找标题
            title = ""
            for shape in slide.Shapes:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    title = shape.TextFrame.TextRange.Text[:80]
                    break
            slides_info.append(f"  第{i}页：{title}({shapes_count}个元素)" if title else f"  第{i}页：{shapes_count}个元素")
        return f"演示文稿：{pres.Name}\n" + "\n".join(slides_info)
    except Exception as e:
        return f"列出失败：{str(e)}"

@mcp.tool()
def ppt_add_text_slide(title: str, content: str = "") -> str:
    """添加一页带标题和内容的幻灯片
    
    Args:
        title: 幻灯片标题
        content: 正文内容
    """
    _init_com()
    import win32com.client
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        pres = ppt.ActivePresentation
        if pres is None:
            return "无活动演示文稿"
        slide = pres.Slides.Add(pres.Slides.Count + 1, 1)  # 标题+内容版式
        slide.Shapes(1).TextFrame.TextRange.Text = title
        if content:
            slide.Shapes(2).TextFrame.TextRange.Text = content
        return f"已添加第{pres.Slides.Count}页：{title}"
    except Exception as e:
        return f"添加失败：{str(e)}"

@mcp.tool()
def ppt_save(filepath: str = "") -> str:
    """保存当前 PPT"""
    _init_com()
    import win32com.client
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        pres = ppt.ActivePresentation
        if pres is None:
            return "无活动演示文稿"
        if filepath:
            pres.SaveAs(filepath)
        else:
            pres.Save()
        return f"已保存：{filepath or pres.FullName}"
    except Exception as e:
        return f"保存失败：{str(e)}"


if __name__ == "__main__":
    mcp.run()
