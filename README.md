# 暨南大学硕士论文格式自查与排版

可独立安装的 Codex skill，依据暨南大学硕士论文模板 2022-4.9、管理学院 22 项格式自查要求，以及用户补充确认的格式标准，检查和统一论文排版。

## 包含内容

- `SKILL.md`：技能入口、操作流程和交付要求。
- `references/standards.md`：字体、标题、页面、目录、页眉页码、图表及参考文献标准。
- `references/checklist.md`：22 项自查映射。
- `references/front-matter.md`：封面字段与声明正文。
- `scripts/audit_docx.py`：只读 DOCX 检查脚本，使用 Python 3 标准库。

最终论文全文字体使用 RGB `000000`，包括关联字符样式，并清除字体主题色覆盖。具体执行标准见技能包内参考文件；使用者当前的明确要求优先。

## 安装与使用

将仓库文件放入 `~/.codex/skills/jnu-thesis-format-check/`，确保该目录下直接包含 `SKILL.md`。

在 Codex 中调用：

```text
$jnu-thesis-format-check 检查并统一这份论文的格式，全文黑色字体，保留正文、数据及参考文献。
```

也可单独运行检查脚本：

```bash
python scripts/audit_docx.py thesis.docx --out audit.json
```

脚本提取可机械验证的格式证据；分页、题注位置、图片内文字等仍需通过 Word、WPS 或 LibreOffice 渲染核对。
