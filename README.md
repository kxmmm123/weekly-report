# Weekly Report Generator

研究生周报生成器 — 基于 Claude Code Skill，自动搜索学术论文并生成格式化的 `.docx` 周报。

## 功能

- **自动论文搜索**：通过 Semantic Scholar API 按关键词搜索真实论文，按引用量排序筛选
- **AI 生成阅读笔记**：基于论文标题和摘要自动生成中文研究总结与阅读笔记
- **项目进展润色**：将简短的自然语言描述扩写为专业化的项目进展条目
- **格式化 Word 文档**：使用 python-docx 生成符合学术规范的 `.docx` 文件（宋体/华文宋体、1.5倍行距、标准页边距）

## 使用方式

在 Claude Code 中输入：

```
/weekly-report
```

按引导依次输入：
1. 姓名
2. 日期范围（默认本周）
3. 研究关键词（逗号分隔，如 `雷达通信一体化, MIMO, 波形设计`）
4. 论文数量（1-10 篇）
5. 项目进展描述（简短即可，AI 会润色）

## 项目结构

```
├── .claude/commands/weekly-report.md   # Skill 定义（交互引导流程）
├── weekly_report_generator.py          # Python 脚本（JSON → .docx）
├── docs/
│   └── superpowers/
│       ├── specs/                      # 设计规格文档
│       └── plans/                      # 实现计划文档
└── .gitignore
```

## 依赖

- Python 3
- [python-docx](https://python-docx.readthedocs.io/)

```bash
pip3 install python-docx
```

## 技术细节

- **论文搜索**：Semantic Scholar Graph API，支持限流重试与 WebSearch 回退
- **排序策略**：优先引用量 > 50 或近 3 年发表的论文
- **去重**：跨关键词精确标题匹配去重
- **文档排版**：标题二号(22pt)，正文小四(12pt)，章节标题四号(14pt)，首行缩进 0.74cm
- **字体**：macOS 使用 STSong，其他平台使用 SimSun

## 输出示例

生成的 `.docx` 包含四个章节：

1. **本周工作总结** — 概述论文阅读与项目进展
2. **文献阅读** — 每篇论文的标题、作者、年份、引用数、研究总结、阅读笔记
3. **项目进展** — 润色后的进展条目
4. **文献阅读总结** — 跨论文的主题关联与洞察
