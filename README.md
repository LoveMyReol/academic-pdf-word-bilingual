# Academic PDF to Word Bilingual QA

这是一个用于学术 PDF、页面截图与可编辑 Word 转换的 Codex Skill。

它将下列事项设为独立的质量门：

- 可编辑 Word 公式：使用 OMML，并优先核验 Word 实际显示，防止积分、求和等大算符出现虚线空框或异常空格。
- 参考文献：从原始书目页建立源清单，保证单一书目标题、连续编号、原文检索字段和原有强调格式。
- 中文译文：按“信、达、雅”审校数学逻辑、术语一致性、中文通顺度与规范的学术文体。

## 安装到仓库

此仓库已采用 Codex 的仓库级目录结构：

    .agents/skills/academic-pdf-word-bilingual/

将仓库克隆到本机后，在该仓库或其子目录中启动 Codex。也可将该技能目录复制到个人技能目录以供所有项目使用。

## 使用

在 Codex 中输入：

    $academic-pdf-word-bilingual

或直接描述包含学术 PDF 转 Word、OMML 公式、参考文献核验、中文译文审校的任务。

## 结构检查

检查脚本不依赖第三方 Python 包：

    py .agents/skills/academic-pdf-word-bilingual/scripts/inspect_docx_academic.py "C:\path\to\document.docx" --require-valid-zip --fail-on-empty-nary

它用于发现 DOCX 结构性风险；公式的真实显示效果仍必须在 Word 或可靠 Office 渲染结果中检查。

## 内容说明

本仓库仅包含 Skill 指令、参考质量门和检查脚本，不包含任何原始 PDF、DOCX、截图、译文或个人凭据。
