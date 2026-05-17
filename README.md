# 📎 Office MCP Server

<p align="center">
  <img src="https://img.shields.io/badge/MCP-Protocol-blue?style=flat-square" alt="MCP">
  <img src="https://img.shields.io/badge/Python-3.10+-green?style=flat-square" alt="Python">
  <img src="https://img.shields.io/badge/Excel-Word-PPT-orange?style=flat-square" alt="Office">
  <img src="https://img.shields.io/badge/License-MIT-yellow?style=flat-square" alt="License">
[![SafeSkill 50/100](https://img.shields.io/badge/SafeSkill-50%2F100_Use%20with%20Caution-orange)](https://safeskill.dev/scan/wade20250715-office-mcp-server)
</p>

**MCP Server for Microsoft Office automation.**  
Control Excel, Word, and PowerPoint through AI agents — read/write cells, insert text, add slides — via standard MCP tools. One server, 13 tools, zero bloat.

> 🎯 **Why this exists**: Office is the most used productivity suite on Earth, yet the MCP ecosystem has virtually no Office tools. Meanwhile, win32com has been capable of controlling Office for decades — it just needed an MCP wrapper.

---

## Architecture

```
┌───────────┐  MCP/JSON-RPC  ┌─────────────┐  COM/win32com  ┌───────────────┐
│ AI Agent  │◄──────────────►│ MCP Server   │◄──────────────►│ Excel / Word  │
│           │   (stdio)      │ (Python 200LOC)│              │ / PowerPoint  │
└───────────┘                └─────────────┘                └───────────────┘
```

Single Python file. No external dependencies beyond `mcp` and `pywin32`. All three Office apps share one MCP server process.

---

## Quick Start

### Prerequisites
- Windows 10/11
- Microsoft Office (tested with Office 2024 LTSC)
- Python 3.10+

### Install

```bash
git clone https://github.com/wade20250715/office-mcp-server.git
cd office-mcp-server
pip install mcp pywin32
```

### Configure Your AI Agent

```json
{
  "mcpServers": {
    "office": {
      "command": "python",
      "args": ["-I", "path/to/office-mcp-server/server.py"],
      "env": { "PYTHONPATH": "" }
    }
  }
}
```

---

## Tools

### Excel (5 tools)

| Tool | Description | Example |
|------|-------------|---------|
| `excel_open` | Open .xlsx file, return sheet names | `excel_open("D:/data/report.xlsx")` |
| `excel_read` | Read cell range as table | `excel_read("Sheet1", "A1", "D20")` |
| `excel_write` | Write value to cell | `excel_write("Sheet1", "B3", "42")` |
| `excel_get_info` | Get workbook overview (sheets, rows, cols) | `excel_get_info()` |
| `excel_save` | Save workbook (optionally to new path) | `excel_save("D:/data/v2.xlsx")` |

### Word (4 tools)

| Tool | Description | Example |
|------|-------------|---------|
| `word_open` | Open .docx, return paragraph count | `word_open("D:/docs/draft.docx")` |
| `word_read_text` | Read full document text | `word_read_text()` |
| `word_insert_text` | Insert text at start/end | `word_insert_text("New paragraph", "end")` |
| `word_save` | Save document | `word_save()` |

### PowerPoint (4 tools)

| Tool | Description | Example |
|------|-------------|---------|
| `ppt_open` | Open .pptx, return slide count | `ppt_open("D:/slides/deck.pptx")` |
| `ppt_list_slides` | List all slides with titles | `ppt_list_slides()` |
| `ppt_add_text_slide` | Add title+content slide | `ppt_add_text_slide("Summary", "Key findings...")` |
| `ppt_save` | Save presentation | `ppt_save()` |

### Example: AI Agent Conversation

```
User: "Open the Q1 report, find the total in cell D20, and create a summary slide"

Agent → excel_open("Q1_report.xlsx")
Agent → excel_read("Sheet1", "D20", "D20")
  → "42420.5"

Agent → ppt_open("summary.pptx")  
Agent → ppt_add_text_slide("Q1 Results", "Total revenue: $42,420.50")
  → "Added slide 5: Q1 Results"
Agent → ppt_save()
```

---

## Verified

```
✓ JSON-RPC initialize handshake
✓ tools/list returns 13 registered tools (5 Excel + 4 Word + 4 PPT)
✓ Server starts in <1 second
✓ All tools importable without Office running (graceful error on actual call)
```

---

## Technical Notes

- `pythoncom.CoInitialize()` called per-tool for COM thread safety
- Office apps start with `Visible = False` (background mode)
- Uses ActiveDocument/ActiveWorkbook/ActivePresentation (opens file first)
- `-I` flag recommended to avoid PYTHONPATH conflicts with mcp

---

## Roadmap

- [ ] Excel: `excel_create_chart`, `excel_apply_formula`, `excel_export_csv`
- [ ] Word: `word_set_style`, `word_find_replace`, `word_export_pdf`
- [ ] PPT: `ppt_add_image_slide`, `ppt_apply_template`, `ppt_export_images`
- [ ] Cross-platform support via python-docx/openpyxl fallback
- [ ] Google Workspace MCP integration

---

## Related Projects

- [autocad-mcp-server](https://github.com/wade20250715/autocad-mcp-server) — AutoCAD MCP
- [google-workspace-mcp](https://github.com/taylorwilsdon/google_workspace_mcp) — Google Workspace MCP (2.4k stars)
- [blender-mcp](https://github.com/ahujasid/blender-mcp) — Blender MCP

---

## License

MIT © [wade20250715](https://github.com/wade20250715)
