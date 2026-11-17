# DOCX Killer

[English](README.md) | 简体中文

适用于 **Claude Code、Codex 和 DSH** 的 Word 文档处理技能，结合 **python-docx**、**docxtpl**、**Pandoc**、**npm Mammoth** 和有针对性的 OOXML 编辑，并明确检查文档结构与排版。

支持文档检查、结构编辑指导、跨文本片段的保守替换、严格模板填充、语义转换和文档包比较。不承诺任意编辑都能无损保留格式，也不自行实现 Word 排版引擎。

## 安装技能

克隆仓库后运行安装脚本。脚本仅使用 Python 标准库，默认将本地技能安装到三个 agent：

```powershell
git clone https://github.com/abigaha/docx-killer.git
cd docx-killer
python install.py
```

也可以单独选择一个或多个目标：

```powershell
python install.py --agent codex
python install.py --agent claude --agent dsh
python install.py --help
```

| Agent | 默认安装位置 |
|---|---|
| Claude Code（`claude`） | `~/.claude/skills/docx-killer` |
| Codex（`codex`） | `~/.codex/skills/docx-killer`；设置 `CODEX_HOME` 时使用其下的 `skills/docx-killer` |
| DSH（`dsh`） | `~/.dsh/skills/docx-killer` |

脚本根据自身位置读取仓库中的技能，因此也可从其他工作目录运行。内容相同的安装会跳过；更新已有安装前，会将原目录完整备份到对应 agent 配置目录下的 `skill-backups` 文件夹。备份位于 `skills` 目录之外，避免被当作另一个技能加载，具体路径会显示在终端中。DSH 已有的自定义描述会保留，技能正文和配套文件更新为仓库版本。DSH 描述支持单行或缩进的 `>`、`|` 块；遇到跨行引号描述时会拒绝更新，保留现有安装，可先将其改为块格式。缓存文件不会被复制。

脚本负责复制技能，依赖需按下一节另行安装。安装后开启新的 agent 会话，使其加载更新后的技能，再提出对应的 DOCX 任务。仓库不包含个人文档或凭据。

## 依赖

- Python，以及 `python-docx`、`docxtpl`、`lxml` 和 `Jinja2`。兼容版本范围随技能发布，位于 `skills/docx-killer/requirements.txt`；仓库根目录的 `requirements.txt` 引用该文件。
- Pandoc 可执行程序。
- Node.js 和 Mammoth，可运行 `npm install -g mammoth`。使用项目本地包时，将 `MAMMOTH_MODULE` 设置为 Mammoth 包目录的绝对路径；Node 按技能脚本所在目录解析模块，而不是当前工作目录。
- 需要检查实际排版时，须有可用的 Office 应用；上述四种工具本身不能渲染 Word 页面。
- 如需 PDF 元数据和逐页预览，还需 Poppler 的 `pdfinfo`、`pdftoppm`。

```powershell
python -m pip install -r requirements.txt
python skills/docx-killer/scripts/doctor.py
python skills/docx-killer/scripts/docx_ops.py --help
```

安装后的技能目录中也包含 `requirements.txt`，可用于独立安装其 Python 依赖。

## 使用示例

```powershell
python skills/docx-killer/scripts/docx_ops.py inspect input.docx --text
python skills/docx-killer/scripts/docx_ops.py inspect input.docx --summary --text --story word/document.xml --paragraph-range 0:20 --tables --media
python skills/docx-killer/scripts/docx_ops.py replace input.docx edited.docx --old 'Old name' --new 'New name' --expected 2
python skills/docx-killer/scripts/docx_ops.py render template.docx context.json generated.docx
python skills/docx-killer/scripts/docx_ops.py compare input.docx edited.docx
python skills/docx-killer/scripts/docx_ops.py verify generated.docx --render auto --output-dir render-output
node skills/docx-killer/scripts/mammoth_extract.cjs input.docx preview.html
```

对新建或修改后的 DOCX，`verify` 检查文档包结构，并用 `python-docx` 重新打开文档。只读检查以及 HTML、Markdown、文本转换，应核对实际发现或输出文件。

存在可用的 Office 渲染器时，`verify` 还会导出 PDF；安装 Poppler 后会为每页生成 PNG 预览。每次渲染尝试使用独立的空工作目录，已有目标 PDF 不会被覆盖。在 Windows 上，`--render auto` 依次尝试 `KWPS.Application`、`Word.Application` 和 LibreOffice，并记录实际使用的 COM ProgID。`--render word` 只尝试 `Word.Application`；使用 WPS 时选择 `--render auto`。

报告区分跳过、不可用、失败和成功导出，并保留回退前的失败原因。PDF 导出成功不等于排版正确，还需查看逐页预览，检查裁切、重叠、分页及字体。

仅使用可信模板，将填充数据作为 JSON 传入。模板填充会拒绝缺失变量，并转义 XML 特殊字符。Mammoth 生成语义 HTML，不复现原文档的视觉排版，也不会自动清理危险 HTML。

## 验证

```powershell
python -m unittest discover -s tests -v
```

集成测试需要上述四种工具均已安装，测试文档会在临时目录中合成。另已使用真实 WPS Office 12.1.0.29690 导出一份两节、两页文档，并检查逐页图像、中文、表格、图片、页眉页脚、脚注和尾注。这证明了本机 WPS 流程可用，不代表所有文档或独立 Microsoft Word 安装都已验证。

工具选择见[技能说明](skills/docx-killer/SKILL.md)，格式保真与检查边界见[验证指南](skills/docx-killer/references/verification.md)，实际执行结果见[验证记录](docs/validation.md)。
