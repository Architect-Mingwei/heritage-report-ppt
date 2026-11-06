# heritage-report-ppt

用于文化遗产、历史、园林、建筑与数字化研究学术汇报的本地技能。支持从论文、长摘要、研究笔记、现有简报、数据和图片出发，完成内容规划、逐页故事线、图表与示意图补充、PPTX 制作及 PDF 验收。

技能定义见 [SKILL.md](SKILL.md)，详细规则位于 `references/`，几何与可读性检查脚本位于 `scripts/audit_pptx.py`。

## 安装

将仓库克隆到本地技能目录：

```powershell
git clone https://github.com/Architect-Mingwei/heritage-report-ppt.git "$env:USERPROFILE\.codex\skills\heritage-report-ppt"
```

安装 Python 依赖：

```powershell
python -m pip install -r requirements.txt
```

## 使用

在对话中调用：

```text
$heritage-report-ppt
```

随后提供论文、长摘要、研究笔记、现有 PPT、数据或图片，并说明汇报时长、听众和交付格式。未提供长摘要时，也可以基于现有素材开始。

## 检查简报

```powershell
python scripts/audit_pptx.py "C:\path\to\report.pptx" --pages 17 -o "C:\path\to\audit.json"
```

检查脚本只报告页面尺寸、越界、文字过小、图片重叠等几何与可读性警告，不替代逐页渲染和人工视觉验收。

## 许可

本仓库未附带开源许可证。如需复制、改编或再分发，请先取得作者许可。
