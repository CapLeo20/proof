---
title: 期刊校样（Author's Proof）批注
status: done
created: 2026-10-04
issue:
tags: [pdf, 论文]
---

# 期刊校样（Author's Proof）批注

## 目标

把作者对编辑部 Q1–Q30 问询的回复，以及额外发现的排版错误（E1–E7），
以标准 PDF 批注的形式标到校样上，每条批注注明页码和行号，便于上传到 Production Forum。

## 输入

原始材料**没有提交到仓库**（当时在会话中直接提供）：

- 原始校样 PDF（Frontiers Author's Proof，29 页，带页边行号）
- `Author_Queries_Response.docx`：作者对问询的 Word 回复

## 期望产出

- 带批注的校样 PDF（不改正文、保留原有加密设置）
- 一份核对清单，列出每条批注和对 Word 回复所做的改动

## 进展 / 结果

已完成（原分支 `claude/pdf-proofreading-line-numbers-c3mipm`，2026-10-04）。

**产出**

| 文件 | 说明 |
|---|---|
| `output/Authors-Proof_annotated.pdf` | 提交用的批注版校样（增量保存，正文和安全设置不变） |
| `output/annotation_checklist.md` | 全部批注清单、对 Word 回复的修改说明、E1–E7 编号对照 |
| `tools/annotate_proof.py` | 读取 Word 回复并在 PDF 上添加批注 |
| `tools/proofmap.py` | 把 PDF 文本位置映射到页边打印的行号 |

**如何复现**（需要把两份原始材料放进 `input/`）：

```bash
cd tasks/2026-10-04-proof-annotation
uv run tools/annotate_proof.py input/proof.pdf input/Author_Queries_Response.docx \
    output/Authors-Proof_annotated.pdf output/annotation_checklist.md
```

**遗留**：Q14 新增的 Table 5 标题措辞需要作者确认（见清单第 1 节）。
