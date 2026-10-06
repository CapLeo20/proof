# Claude Code 工作台

这个仓库用来存放交给 **Claude Code（云端，claude.ai/code）** 处理的各种零碎任务。
它不按任务类型划分，而是**一个任务一个文件夹**：写文档、处理 PDF、跑脚本、整理数据都可以放在这里。

## 目录结构

```
.
├── CLAUDE.md                 # 给 Claude 的工作约定（每次会话自动读取）
├── README.md                 # 本文件，给人看的说明
├── scripts/tasks.py          # 新建任务 / 列出所有任务
├── .github/                  # Issue 模板、PR 模板
└── tasks/
    ├── _template/            # 新任务模板
    └── 2026-10-04-proof-annotation/   # 每个任务一个文件夹：<日期>-<英文短名>
        ├── README.md         # 任务卡：目标、输入、产出、状态、结果
        ├── input/            # 原始材料（你上传的文件）
        ├── output/           # 交付物
        └── tools/ 或 *.py    # 这个任务专用的脚本
```

## 怎么给 Claude 派任务

**方式 A：直接在会话里说（最简单）**

在 claude.ai/code 选中这个仓库，然后输入，例如：

> 新任务：把 input 里的 Excel 按月份汇总，输出一个 Markdown 表格。

Claude 会按 `CLAUDE.md` 的约定新建 `tasks/<日期>-<短名>/`，把需求写进任务卡，然后开始工作。

**方式 B：先开 Issue，再派给 Claude（顺便练习 GitHub）**

1. 在 GitHub 上点 **Issues → New issue → 新任务**，按模板填写。
2. 在 claude.ai/code 里说：`处理 issue #3`。
3. Claude 完成后提 PR，PR 描述里写 `Closes #3`，合并后 Issue 会自动关闭。

**继续一个旧任务**

> 继续 tasks/2026-10-04-proof-annotation：……

**需要上传文件时**

云端会话每次都是从 GitHub 重新克隆仓库，所以文件要**先放进仓库**：
在 GitHub 网页上进入 `tasks/<任务名>/input/`，点 **Add file → Upload files**，提交后再开会话。

## 一次任务的完整流程

```
派任务 → Claude 在自己的 claude/xxx 分支上工作并推送
       → 你在 GitHub 上开 PR（或让 Claude 开）→ 检查改动 → Merge 到 main
       → 删除这个分支
```

每个云端会话都会用一个独立的 `claude/...` 分支，所以多个任务可以同时进行，互不干扰。

## 隔离规则（摘要）

| 层级 | 隔离的是什么 | 做法 |
|---|---|---|
| 文件夹 | 文件 | 每个任务只改自己的 `tasks/<任务名>/`，不碰别的任务和根目录 |
| 分支 / PR | 改动 | 一个会话一个分支，通过 PR 合并到 `main` |
| 依赖 | Python 包版本 | 脚本头部写 PEP 723 依赖声明，用 `uv run 脚本.py` 运行，不装全局依赖 |
| 上下文 | Claude 的注意力 | 一个会话只做一个任务；任务的背景写在任务卡里，不靠聊天记录 |
| 云端环境 | 网络、密钥 | 需要 API Key 或特殊网络的任务，单独建一个 cloud environment |
| 仓库 | 可见性、生命周期 | 敏感材料放私有仓库；长期维护的项目"毕业"成独立仓库 |

详细约定见 [`CLAUDE.md`](CLAUDE.md)。

## 常用命令

```bash
python3 scripts/tasks.py list                         # 列出所有任务及状态
python3 scripts/tasks.py new excel-summary "Excel 按月汇总"  # 新建任务文件夹
uv run tasks/<任务名>/xxx.py                           # 运行任务脚本（自动装好依赖）
```
