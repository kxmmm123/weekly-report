# Weekly Report Generator Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Claude Code slash command + Python script that generates formatted graduate student weekly reports (.docx) with real paper search from Semantic Scholar API.

**Architecture:** Two-component system — a `.claude/commands/weekly-report.md` slash command handles interactive data collection, paper search via Semantic Scholar API (curl), and AI-generated notes; a `weekly_report_generator.py` Python script reads a JSON intermediate file and produces a formatted .docx output.

**Tech Stack:** Claude Code slash commands, Semantic Scholar API (free, no key), Python 3 + python-docx, Bash/curl

---

### Task 1: Create the Python script — weekly_report_generator.py

**Files:**
- Create: `weekly_report_generator.py`

This script reads a JSON file and produces a .docx file. It has no external dependencies beyond python-docx and the standard library.

- [ ] **Step 1: Write weekly_report_generator.py**

```python
#!/usr/bin/env python3
"""Generate a formatted weekly report (.docx) from JSON data."""

import json
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    print("Error: python-docx is not installed. Run: pip3 install python-docx")
    sys.exit(1)

FONT_NAME = "SimSun"
FONT_FALLBACK = "STSong"


def get_font_name():
    """Return the first available CJK font."""
    # On macOS, SimSun may not exist; STSong is the fallback
    return FONT_NAME


def set_run_font(run, size_pt, bold=False):
    """Apply font settings to a run."""
    font = get_font_name()
    run.font.name = font
    run.font.size = Pt(size_pt)
    run.bold = bold
    # CJK font requires eastAsia setting
    from docx.oxml.ns import qn
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font)


def add_title(doc, text):
    """Add centered bold title (二号 = 22pt)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    set_run_font(run, 22, bold=True)


def add_centered_text(doc, text):
    """Add centered text (小四 = 12pt)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    set_run_font(run, 12)


def add_heading_text(doc, text):
    """Add section heading (四号 = 14pt, bold)."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, 14, bold=True)


def add_body_text(doc, text, indent=False):
    """Add body text (小四 = 12pt), optionally indented."""
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.74)  # ~2 chars
    run = p.add_run(text)
    set_run_font(run, 12)


def add_paper_entry(doc, idx, paper):
    """Add a single paper entry."""
    title = paper.get("title", "Unknown Title")
    authors = paper.get("authors", "Unknown")
    year = paper.get("year", "N/A")
    citations = paper.get("citations", 0)
    summary = paper.get("research_summary", "")
    notes = paper.get("reading_notes", "")

    # Paper title line
    p = doc.add_paragraph()
    run = p.add_run(f"论文 {idx}：{title}")
    set_run_font(run, 12, bold=True)

    # Meta line
    add_body_text(doc, f"作者：{authors} | 年份：{year} | 引用数：{citations}")

    # Research summary
    add_body_text(doc, f"研究总结：{summary}", indent=True)

    # Reading notes
    add_body_text(doc, f"阅读笔记：{notes}", indent=True)

    # Blank line separator
    doc.add_paragraph()


def generate_report(data, output_path):
    """Generate the .docx report from JSON data."""
    doc = Document()

    # Set default font for the document
    style = doc.styles['Normal']
    style.font.name = FONT_NAME
    style.font.size = Pt(12)
    style.paragraph_format.line_spacing = 1.5

    # Page margins: top/bottom 2.54cm, left/right 3.18cm
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.18)
        section.right_margin = Cm(3.18)

    # Title
    add_title(doc, data.get("title", "研究生周报"))

    # Date range
    date_range = data.get("date_range", {})
    start = date_range.get("start", "")
    end = date_range.get("end", "")
    add_centered_text(doc, f"时间：{start} - {end}")

    # Name
    add_centered_text(doc, f"姓名：{data.get('name', '')}")

    # Blank line
    doc.add_paragraph()

    # Section 1: Summary
    add_heading_text(doc, "一、本周工作总结")
    add_body_text(doc, data.get("summary", ""), indent=True)

    # Section 2: Papers
    add_heading_text(doc, "二、文献阅读")
    papers = data.get("papers", [])
    for idx, paper in enumerate(papers, 1):
        add_paper_entry(doc, idx, paper)

    # Section 3: Project progress
    add_heading_text(doc, "三、项目进展")
    for item in data.get("project_progress", []):
        add_body_text(doc, f"• {item}")

    # Section 4: Paper summary
    add_heading_text(doc, "四、文献阅读总结")
    add_body_text(doc, data.get("paper_summary", ""), indent=True)

    doc.save(output_path)
    return output_path


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 weekly_report_generator.py <json_path> [output_path]")
        sys.exit(1)

    json_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "研究生周报.docx"

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    result = generate_report(data, output_path)
    print(f"Report generated: {result}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Test the script with sample data**

Create a test JSON file:

```bash
cat > /tmp/test_report_data.json << 'EOF'
{
  "title": "研究生周报",
  "date_range": {
    "start": "2026-05-25",
    "end": "2026-05-31"
  },
  "name": "测试用户",
  "summary": "本周共阅读 2 篇论文，主要涉及大语言模型领域；项目方面完成了数据处理模块的开发。",
  "papers": [
    {
      "title": "Attention Is All You Need",
      "authors": "Ashish Vaswani, Noam Shazeer, Niki Parmar",
      "year": 2017,
      "citations": 120000,
      "research_summary": "提出了Transformer架构，完全基于注意力机制，摒弃了传统的循环和卷积结构。",
      "reading_notes": "Transformer的核心创新在于自注意力机制，能够并行处理序列数据，大幅提升了训练效率。",
      "url": "https://..."
    },
    {
      "title": "BERT: Pre-training of Deep Bidirectional Transformers",
      "authors": "Jacob Devlin, Ming-Wei Chang, Kenton Lee",
      "year": 2019,
      "citations": 90000,
      "research_summary": "提出了BERT模型，通过掩码语言模型和下一句预测任务进行预训练。",
      "reading_notes": "BERT的双向编码特性使其在多项NLP任务上取得了突破性进展，预训练-微调范式影响深远。",
      "url": "https://..."
    }
  ],
  "project_progress": [
    "完成了数据处理模块的开发与单元测试",
    "优化了特征提取流程，效率提升约30%"
  ],
  "paper_summary": "本周阅读的两篇论文均与Transformer架构相关。Vaswani等人的工作奠定了现代NLP的基础架构，而BERT则展示了该架构在预训练范式下的巨大潜力。这些工作对当前项目的模型选型具有重要参考价值。"
}
EOF
python3 weekly_report_generator.py /tmp/test_report_data.json /tmp/测试周报.docx
```

Expected: `Report generated: /tmp/测试周报.docx` and file exists.

- [ ] **Step 3: Verify the generated .docx is valid**

```bash
python3 -c "from docx import Document; d = Document('/tmp/测试周报.docx'); print(f'Paragraphs: {len(d.paragraphs)}'); print(d.paragraphs[0].text)"
```

Expected: Shows paragraph count and title "研究生周报".

- [ ] **Step 4: Clean up test data**

```bash
rm /tmp/test_report_data.json /tmp/测试周报.docx
```

- [ ] **Step 5: Commit**

```bash
git add weekly_report_generator.py
git commit -m "feat: add weekly report docx generator script"
```

---

### Task 2: Create the slash command — weekly-report.md

**Files:**
- Create: `.claude/commands/weekly-report.md`

This is the Claude Code slash command that handles all interaction, paper search, content generation, and script invocation.

- [ ] **Step 1: Create .claude/commands directory**

```bash
mkdir -p .claude/commands
```

- [ ] **Step 2: Write the slash command file**

```markdown
---
name: weekly-report
description: Use when you need to generate a graduate student weekly report (.docx) with literature review and project progress
---

# 研究生周报生成器

You are a graduate student weekly report generator. Follow these steps IN ORDER. Do NOT skip steps. Do NOT proceed without user input for each step.

## Step 1: Collect Name

Ask the user for their name. If you have previously saved the user's name in memory, confirm it with them instead of asking again.

## Step 2: Collect Date Range

Ask the user for the date range. Default is the current week (Monday to Sunday). Present the default and let them confirm or override.

Use this bash command to calculate this week's Monday and Sunday:
```bash
python3 -c "
from datetime import datetime, timedelta
today = datetime.now()
monday = today - timedelta(days=today.weekday())
sunday = monday + timedelta(days=6)
print(f'{monday.strftime(\"%Y-%m-%d\")} to {sunday.strftime(\"%Y-%m-%d\")}')
"
```

## Step 3: Collect Research Keywords

Ask the user for research keywords/domains, separated by commas. Example: "大语言模型, 知识图谱, 检索增强生成"

## Step 4: Collect Paper Count

Ask the user how many papers they want to include. Default is 3.

## Step 5: Collect Project Progress

Ask the user to describe their project progress in natural language. Tell them they can be brief — you will polish and expand it.

## Step 6: Search Papers

For each keyword, search Semantic Scholar API using curl:

```bash
curl -s "https://api.semanticscholar.org/graph/v1/paper/search?query=KEYWORD&limit=LIMIT&fields=title,authors,year,abstract,citationCount,url"
```

Replace KEYWORD with the URL-encoded keyword and LIMIT with the paper count (or slightly more to allow for filtering).

**Filtering strategy:**
- Prefer papers with citationCount > 50 OR year >= 2023
- If a keyword returns no results, try WebSearch as fallback: search for "site:scholar.google.com KEYWORD"
- Deduplicate across keywords by paper title
- Select the top N papers overall

**If a paper has no abstract**, note it as "暂无摘要" and generate notes based on the title only.

## Step 7: Generate Content

Based on the search results, generate the following content:

### 7a: Work Summary
Write 1-2 sentences summarizing: how many papers were read, what domains they cover, and a brief project progress summary.

### 7b: Paper Notes
For each paper, generate:
- **research_summary** (研究总结): 2-3 sentences in Chinese summarizing the paper's contribution and method based on its title and abstract
- **reading_notes** (阅读笔记): 3-5 sentences in Chinese with personal insights, potential applications, and critical thoughts. Write as if the student actually read and reflected on the paper.

Paper titles stay in English. Author names stay as returned by the API (format: "Author1, Author2, Author3", max 3 authors, then "et al.").

### 7c: Project Progress
Polish and expand the user's project description into 2-4 bullet points. Make it sound professional and substantive but not exaggerated. Use technical language appropriate for the research domain.

### 7d: Paper Summary
Write 2-3 paragraphs in Chinese summarizing the themes, connections, and insights across all papers read this week. Highlight how they relate to the student's research direction.

## Step 8: Write JSON and Generate .docx

Assemble all data into a JSON file following this schema:

```json
{
  "title": "研究生周报",
  "date_range": { "start": "YYYY-MM-DD", "end": "YYYY-MM-DD" },
  "name": "...",
  "summary": "...",
  "papers": [
    {
      "title": "...",
      "authors": "...",
      "year": 2024,
      "citations": 123,
      "research_summary": "...",
      "reading_notes": "...",
      "url": "..."
    }
  ],
  "project_progress": ["...", "..."],
  "paper_summary": "..."
}
```

Write the JSON to `/tmp/weekly_report_data.json`:

```bash
cat > /tmp/weekly_report_data.json << 'ENDJSON'
{the JSON content}
ENDJSON
```

Then run the generator script:

```bash
python3 weekly_report_generator.py /tmp/weekly_report_data.json
```

The output file will be saved as `研究生周报.docx` in the current directory.

If python-docx is not installed, prompt the user to run: `pip3 install python-docx`, then retry.

## Step 9: Report Result

Tell the user:
- The report has been generated
- The file path
- A brief summary of what's in the report (paper count, topics covered)

## Error Handling

- Semantic Scholar API returns no results → Fall back to WebSearch for Google Scholar
- API request fails → Retry once. If still fails, tell user to check network
- Keywords too broad → Suggest the user narrow the scope
- Paper missing abstract → Use "暂无摘要", generate notes from title only
```

- [ ] **Step 3: Verify the command is discoverable**

```bash
ls -la .claude/commands/weekly-report.md
```

Expected: File exists.

- [ ] **Step 4: Commit**

```bash
git add .claude/commands/weekly-report.md
git commit -m "feat: add weekly-report slash command"
```

---

### Task 3: End-to-end integration test

**Files:**
- No new files created

- [ ] **Step 1: Verify python-docx is installed**

```bash
python3 -c "import docx; print('python-docx version:', docx.__version__)"
```

If not installed: `pip3 install python-docx`

- [ ] **Step 2: Test the full pipeline manually**

Run the slash command `/weekly-report` in Claude Code and verify:
1. It asks for name
2. It asks for date range (with default)
3. It asks for keywords
4. It asks for paper count
5. It asks for project progress
6. It searches Semantic Scholar API
7. It generates content and JSON
8. It calls the Python script
9. It produces a .docx file

- [ ] **Step 3: Open the generated .docx and verify formatting**

Check that:
- Title "研究生周报" is centered and bold
- Date range and name are centered
- Section headings (一、二、三、四) are bold
- Paper entries have title (bold), metadata, summary, and notes
- Project progress has bullet points
- Line spacing is 1.5x
- Page margins look correct

- [ ] **Step 4: Final commit if any fixes needed**

```bash
git add -A
git commit -m "fix: address integration test issues"
```
