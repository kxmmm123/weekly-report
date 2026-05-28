# 研究生周报生成器 - 设计文档

## 目标

为研究生生成格式化的周报 .docx 文件，核心场景是应付导师检查。周报包含文献阅读和项目进度两部分，论文信息来自真实搜索，项目描述由 AI 润色扩充。

## 架构

```
weekly-report/
├── skill.md                    # Skill 定义：交互流程 + 内容生成 + 调用脚本
└── weekly_report_generator.py  # Python 脚本：读取 JSON → 生成 .docx
```

### 职责划分

| 组件 | 职责 |
|------|------|
| skill.md | 分步引导收集信息；搜索论文；生成笔记内容；写入中间 JSON；调用脚本 |
| weekly_report_generator.py | 读取 JSON 数据；按模板生成 .docx；返回文件路径 |

### 执行流程

1. 用户输入 `/weekly-report`
2. Skill 分步引导收集：
   - 姓名
   - 时间范围（默认本周周一到周日）
   - 研究关键词（逗号分隔）
   - 论文数量
   - 项目进度描述（自然语言）
3. Skill 通过 Bash curl 调用 Semantic Scholar API 搜索论文
4. Skill 基于搜索结果（标题、摘要、作者、年份、引用数）生成中文笔记和总结
5. Skill 将所有内容写入 JSON 中间文件（`/tmp/weekly_report_data.json`）
6. Skill 调用 `python3 weekly_report_generator.py /tmp/weekly_report_data.json`
7. 脚本读取 JSON，生成 .docx，保存到当前目录
8. Skill 告知用户文件路径

## 论文搜索

- **数据源**：Semantic Scholar API（免费，无需 key）
- **搜索策略**：每个关键词分别搜索，按引用数排序，取前 N 篇（去重）
- **API 端点**：`https://api.semanticscholar.org/graph/v1/paper/search`
- **请求参数**：
  - `query`：关键词
  - `limit`：论文数量
  - `fields`：`title,authors,year,abstract,citationCount,url`
- **排序**：优先选高被引或近 3 年的新论文
- **语言**：论文标题保留原文（英文），笔记和总结用中文

## 周报模板

```
研究生周报
时间：2026.05.25 - 2026.05.31
姓名：XXX

一、本周工作总结
  本周共阅读 N 篇论文，主要涉及 [关键词] 领域；项目方面 [一句话总结]。

二、文献阅读

  论文 1：[英文标题]
  作者：[作者列表] | 年份：2024 | 引用数：123
  研究总结：[中文总结 2-3 句]
  阅读笔记：[中文笔记 3-5 句]

  论文 2：...

三、项目进展
  [润色后的项目进度描述，分点列出]

四、文献阅读总结
  [对本周所读论文的整体总结，2-3 段]
```

## JSON 中间文件格式

```json
{
  "title": "研究生周报",
  "date_range": {
    "start": "2026-05-25",
    "end": "2026-05-31"
  },
  "name": "XXX",
  "summary": "本周共阅读 3 篇论文...",
  "papers": [
    {
      "title": "Paper Title",
      "authors": "Author1, Author2, Author3",
      "year": 2024,
      "citations": 123,
      "research_summary": "中文总结...",
      "reading_notes": "中文笔记...",
      "url": "https://..."
    }
  ],
  "project_progress": [
    "完成了 XXX 模块的开发与测试",
    "优化了 YYY 算法的性能，提升约 20%"
  ],
  "paper_summary": "本周阅读的论文主要围绕..."
}
```

## .docx 格式规范

- **标题**："研究生周报"，SimSun（宋体，macOS 回退 STSong），二号字，居中，加粗
- **时间/姓名**：SimSun（macOS 回退 STSong），小四号字，居中
- **一级标题**（一、二、三、四）：SimSun（macOS 回退 STSong），四号字，加粗
- **正文内容**：SimSun（macOS 回退 STSong），小四号字
- **论文标题**：SimSun（macOS 回退 STSong），小四号字，加粗
- **页边距**：上下 2.54cm，左右 3.18cm（Word 默认）
- **行距**：1.5 倍行距

## 依赖

- Python 3 + python-docx
- Semantic Scholar API（免费，curl 调用）
- 脚本启动时检查 python-docx 是否安装，未安装则提示 `pip3 install python-docx`

## Skill 交互流程（skill.md）

```
Step 1: 询问姓名（首次使用时，之后可记忆到 user memory）
Step 2: 询问时间范围（默认本周，格式 YYYY-MM-DD 到 YYYY-MM-DD）
Step 3: 询问研究关键词（逗号分隔，如 "大语言模型, 知识图谱"）
Step 4: 询问论文数量（默认 3）
Step 5: 询问项目进度描述（自然语言，可多行）
Step 6: 搜索论文 + 生成内容 + 生成 .docx
Step 7: 告知文件路径
```

## 边界情况

- Semantic Scholar API 无结果：降级用 WebSearch 搜索 Google Scholar
- API 请求失败：重试一次，再失败则报错让用户检查网络
- 关键词过泛（如 "AI"）：提示用户缩小范围
- 论文无摘要：标注"暂无摘要"，仅基于标题生成笔记
