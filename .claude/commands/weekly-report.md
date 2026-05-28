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
from datetime import date, timedelta
today = date.today()
monday = today - timedelta(days=today.weekday())
sunday = monday + timedelta(days=6)
print(f'{monday} to {sunday}')
"
```

## Step 3: Collect Research Keywords

Ask the user for research keywords/domains, separated by commas. Example: "大语言模型, 知识图谱, 检索增强生成"

## Step 4: Collect Paper Count

Ask the user how many papers they want to include. Default is 3. Validate the input is a positive integer (1-10). If invalid, ask again.

## Step 5: Collect Project Progress

Ask the user to describe their project progress in natural language. Tell them they can be brief — you will polish and expand it.

## Step 6: Search Papers

Before searching, verify the generator script exists:
```bash
ls weekly_report_generator.py
```
If not found, tell the user the script is missing and cannot proceed.

For each keyword, search Semantic Scholar API using curl. URL-encode keywords using Python:
```bash
python3 -c "import urllib.parse; print(urllib.parse.quote('KEYWORD'))"
```
Then search:
```bash
curl -s "https://api.semanticscholar.org/graph/v1/paper/search?query=ENCODED_KEYWORD&limit=LIMIT&fields=title,authors,year,abstract,citationCount,url"
```

Set LIMIT to 2x the desired paper count to allow for filtering.

**Ranking strategy:** Sort by citationCount descending, prefer papers with citationCount > 50 OR year >= 2023.

**Filtering strategy:**
- Deduplicate across keywords by exact paper title match
- Select the top N papers by citationCount
- If a keyword returns no results, try WebSearch as fallback: search for "site:scholar.google.com KEYWORD"
- If ALL keywords return no results even after fallback, tell the user and suggest different keywords

**If a paper has no abstract**, note it as "暂无摘要" and generate notes based on the title only.

**API rate limiting:** If you receive HTTP 429, wait 5 seconds and retry once.

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

**IMPORTANT: JSON escaping** — All string values must be valid JSON. Escape double quotes as `\"`, backslashes as `\\`, and newlines as `\n`. Use Python to write the file safely instead of a heredoc:

```bash
python3 -c "
import json
data = {the assembled dict}
with open('/tmp/weekly_report_data.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
print('JSON written successfully')
"
```

Then verify the JSON is valid:
```bash
python3 -c "import json; json.load(open('/tmp/weekly_report_data.json')); print('JSON valid')"
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
- All search methods fail → Tell user, suggest different keywords, and stop
- API request fails → Retry once (with 5s delay if rate-limited). If still fails, tell user to check network
- Keywords too broad → Suggest the user narrow the scope
- Paper missing abstract → Use "暂无摘要", generate notes from title only
- python-docx not installed → Prompt user to install with pip3
- JSON write fails → Check file permissions on /tmp
- Generator script missing → Tell user the script file is not found
