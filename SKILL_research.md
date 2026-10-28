---
name: topic-research-brief
description: Multi-source topic research aggregator that searches news,
  encyclopedia entries, statistical data, academic papers, industry reports, and
  news facts/typical cases based on user-provided keywords, then categorizes and
  synthesizes findings into a formatted Word/PDF research brief with cover page,
  table of contents, data tables, and source citations. The 'news facts &
  typical cases' category finds real-world examples, specific incidents,
  historical precedents, and related social phenomena — e.g. searching 网络暴力
  surfaces 刘翔's case; searching 社会达尔文主义 finds 纳粹德国 and 末位淘汰制. This skill should
  be used when the user provides a research topic or keywords and asks to
  search, gather, collect, or integrate information from multiple sources for
  academic assignments, debate preparation, or research projects. Triggers
  include phrases like 搜集资料, 整合研究, 帮我查, 搜一下这个话题, 研究简报, 选题调研, 辩论备赛, 收集素材, topic
  research, multi-source search.
agent_created: true
disable: false
---

# Topic Research Brief

## Overview

Search across six source categories (news, encyclopedia, statistics, academic papers,
industry reports, news facts & typical cases) based on a user-provided topic, then
categorize and synthesize all findings into a professionally formatted research brief
document (Word/PDF) with cover page, table of contents, categorized chapters, data
tables, and source citations.

Designed for university journalism/communication assignments and debate preparation.

## When to Use

Trigger this skill when the user:

- Provides a topic or set of keywords and asks to search/gather/collect information
- Needs multi-source research for an assignment, paper, or debate
- Asks for a "research brief" or "research summary" on a topic
- Wants categorized findings from different types of sources (news, data, papers, etc.)
- Says "帮我查资料" / "搜集素材" / "选题调研" / "辩论备赛" / "整合研究"

## Workflow

### Step 1: Clarify the Research Scope

Before searching, confirm or infer the following from the user's input:

- **Topic**: the main subject of research
- **Keywords**: specific terms to search for (ask the user if not clear)
- **Context**: assignment, debate prep, or other use case (affects depth and focus)
- **Output format**: Word (.docx) or PDF (default: docx if not specified)

If the user's request is vague, ask 1-2 focused questions to narrow the scope. Do not
over-interrogate — make reasonable assumptions and proceed.

### Step 2: Multi-Source Search

Search across six categories using `WebSearch`. For each category, craft tailored
queries. Refer to `references/source_catalog.md` for detailed search strategies,
recommended sources, and query templates.

**Use `query_keyword_groups`** in WebSearch to search multiple angles simultaneously
(max 5 groups per call). This is more efficient than separate searches.

**Search plan (run as parallel WebSearch calls where possible):**

1. **News** — `WebSearch(query='"{topic}" 最新 新闻', topic="news")`
2. **Encyclopedia** — `WebSearch(query='"{topic}" 百科 定义 概念')`
3. **Statistics** — `WebSearch(query='"{topic}" 统计数据 报告', topic="general")`
4. **Academic papers** — `WebSearch(query='"{topic}" 论文 研究', topic="academic")`
5. **Industry reports** — `WebSearch(query='"{topic}" 行业报告 白皮书 智库')`
6. **News facts & typical cases** — Multi-angle searches for real-world examples,
   specific incidents, and related social/historical phenomena. Use 2-3 WebSearch
   calls with `query_keyword_groups` to maximize coverage. Refer to
   `references/source_catalog.md` Category 6 for detailed strategies.

   Example search angles:
   - `"{topic}" 案例 事件 人物 经历` — specific real-world cases and people
   - `"{topic}" 典型事件 真实故事 事例` — notable events and personal stories
   - `"{topic}" 社会现象 相关事件 联系` — broader social phenomena linked to the topic
   - `"{topic}" 历史 先例 类似 类比` — historical precedents and analogous cases
   - `"{topic}" 现实生活中 表现 体现` — how the topic manifests in real life

   **Key principle**: Think broadly about what real-world examples, historical events,
   social phenomena, and institutional practices relate to the topic. For instance:
   - Searching "网络暴力" should surface specific cases like 刘翔遭受网络暴力,
     崔雪莉事件, etc.
   - Searching "社会达尔文主义" should find 纳粹种族政策, 企业末位淘汰制,
     社会阶层固化, etc.
   - Searching "消费主义" might find 钻石营销史, 盲盒经济, 双十一现象, etc.
   - Searching "信息茧房" should surface 算法推荐争议, 极端化案例, etc.

For each search, collect 5-10 result URLs with titles and snippets.

### Step 3: Deep Fetch

Select the 2-4 most promising URLs per category. Use `WebFetch` to extract full content.
Prioritize:
- Official/government sources over blogs
- Recent content over older content
- Primary sources over secondary reporting
- Data-rich pages over opinion pieces

For each fetched page, use a targeted prompt:
```
WebFetch(
    url="<url>",
    prompt="Extract: 1) Title 2) Author/Source 3) Date 4) Key facts and data points 5) Main arguments or findings 6) Any statistics or numbers with their context. Return in Chinese."
)
```

### Step 4: Synthesize and Structure

After collecting content from all sources, synthesize the findings:

1. **Write an executive summary** (1-3 paragraphs) that provides an overview of the topic
   based on all gathered sources. This should be original analysis, not copied text.

2. **Identify 3-5 key findings** — the most important data points, conclusions, or insights
   across all sources. These should be concise, self-contained statements with specific data.

3. **Organize items by category** — place each source's extracted content into the
   appropriate category (news, encyclopedia, statistics, papers, reports, cases).

4. **For each item**, write:
   - A 2-3 sentence summary in your own words
   - Key points as a bullet list
   - Specific data points or quotes with attribution
   - Source citation (title, source name, date, URL)

5. **Build a source list** — compile all cited sources with type labels.

### Step 5: Build the JSON Content File

Create a JSON file at a temporary path (e.g., `<workspace>/research_content.json`)
following this schema:

```json
{
    "title": "研究简报标题",
    "subtitle": "副标题（可选）",
    "date": "2024-12-17",
    "topic_keywords": ["关键词1", "关键词2"],
    "context": "大学新传作业 / 辩论备赛",
    "executive_summary": "1-3段研究概述...",
    "key_findings": ["核心发现1", "核心发现2"],
    "categories": [
        {
            "name": "新闻资讯",
            "description": "本章节涵盖的内容简介（1-2句）",
            "items": [
                {
                    "type": "news",
                    "title": "新闻标题",
                    "source": "来源媒体",
                    "date": "2024-12-10",
                    "url": "https://...",
                    "summary": "内容摘要（2-3句）",
                    "key_points": ["要点1", "要点2"],
                    "quote": "重要引用（可选）"
                }
            ]
        },
        {
            "name": "百科诠释",
            "items": [
                {
                    "type": "encyclopedia",
                    "title": "词条标题",
                    "source": "百度百科",
                    "url": "https://...",
                    "definition": "定义内容",
                    "background": "历史背景（可选）",
                    "key_features": ["特征1", "特征2"]
                }
            ]
        },
        {
            "name": "统计数据",
            "items": [
                {
                    "type": "data",
                    "title": "数据标题",
                    "source": "国家统计局",
                    "date": "2024-06",
                    "url": "https://...",
                    "description": "数据说明",
                    "table": {
                        "headers": ["年份", "指标1", "指标2"],
                        "rows": [["2020", "100", "200"]]
                    },
                    "analysis": "数据分析（可选）"
                }
            ]
        },
        {
            "name": "学术论文",
            "items": [
                {
                    "type": "paper",
                    "title": "论文标题",
                    "authors": "作者",
                    "journal": "期刊",
                    "year": "2024",
                    "url": "https://...",
                    "abstract": "论文摘要",
                    "key_findings": ["发现1", "发现2"]
                }
            ]
        },
        {
            "name": "行业报告与智库分析",
            "items": [
                {
                    "type": "report",
                    "title": "报告标题",
                    "organization": "发布机构",
                    "date": "2024-03",
                    "url": "https://...",
                    "summary": "报告摘要",
                    "key_findings": ["发现1", "发现2"]
                }
            ]
        },
        {
            "name": "新闻事实与典型案例",
            "description": "本章节涵盖与主题相关的真实事件、典型案例和社会现象",
            "items": [
                {
                    "type": "case",
                    "title": "案例标题",
                    "source": "来源媒体/网站",
                    "date": "2024-06",
                    "url": "https://...",
                    "case_type": "典型事件/社会现象/历史先例/人物故事",
                    "summary": "案例摘要（2-3句，描述事件/现象的核心内容）",
                    "key_points": ["要点1", "要点2"],
                    "relevance": "与主题的关联说明（为什么这个案例与搜索主题相关）",
                    "quote": "重要引用（可选）"
                }
            ]
        }
    ],
    "sources": [
        {"title": "来源标题", "url": "https://...", "type": "news"}
    ]
}
```

**Item types and required fields:**

| type          | Required fields                                      | Optional fields          |
|---------------|------------------------------------------------------|--------------------------|
| news          | title, source, url, summary                          | date, key_points, quote  |
| encyclopedia  | title, source, url, definition                       | background, key_features |
| data          | title, source, url, description                      | date, table, analysis    |
| paper         | title, url, abstract                                 | authors, journal, year, key_findings |
| report        | title, organization, url, summary                    | date, key_findings       |
| case          | title, url, summary, relevance                       | source, date, case_type, key_points, quote |

**Data tables**: When statistical data is found, structure it as a table object with
`headers` (array of strings) and `rows` (array of arrays). The script renders these as
formatted Word tables with header row shading.

### Step 6: Generate the Research Brief

Run the generation script:

```bash
<C:\Users\DELL\.workbuddy\binaries\python\envs\default\Scripts\python.exe> <skill_dir>/scripts/generate_brief.py <input.json> --format docx --output <output.docx>
```

**Parameters:**
- `--format`: `docx` or `pdf` (PDF requires Microsoft Word installed)
- `--output` / `-o`: Output file path (optional; defaults to input filename with new extension)

**Prerequisites:** The Python venv with dependencies must exist. If not:
```bash
<C:\Users\DELL\.workbuddy\binaries\python\versions\3.13.12\python.exe> -m venv <C:\Users\DELL\.workbuddy\binaries\python\envs\default>
<C:\Users\DELL\.workbuddy\binaries\python\envs\default\Scripts\pip> install python-docx docx2pdf
```

### Step 7: Deliver the Result

After generation:
1. Verify the output file exists and has a non-zero size
2. Present the file via `present_files`
3. Provide a brief summary of what was searched, how many sources were found per category, and key findings
4. Clean up the temporary JSON file
5. **Important**: Remind the user to open the Word document and right-click the table of contents area, then select "Update Field" (更新域) to populate the TOC

## Document Structure

The generated research brief includes:

1. **Cover page** — title, subtitle, date, research context, keywords
2. **Table of contents** — auto-generated TOC field (requires manual update in Word)
3. **Executive summary** — overview paragraph + key findings bullet list
4. **Categorized chapters** — one chapter per source type:
   - News (新闻资讯) — latest developments with summaries and key points
   - Encyclopedia (百科诠释) — definitions, backgrounds, key features
   - Statistics (统计数据) — data tables with analysis
   - Academic papers (学术论文) — abstracts and key findings
   - Industry reports (行业报告与智库分析) — report summaries and findings
   - News facts & typical cases (新闻事实与典型案例) — real-world events, notable
     incidents, social phenomena, and historical precedents linked to the topic
5. **Source list** — complete citation list with type labels and URLs
6. **Page numbers** — centered in footer

## Search Strategy Notes

- **Chinese sources first**: Prioritize Chinese-language sources (新华网, 人民网, 百度百科,
  国家统计局, 知网, 艾瑞咨询, etc.). Supplement with English sources only when the topic
  has significant international dimensions.
- **Data accuracy**: Always note the original data source and publication date for
  statistics. Cross-reference numbers across multiple sources when possible.
- **Recency**: For news, prioritize the last 3-6 months. For academic papers, prioritize
  the last 5 years but include seminal older works. For statistics, note the data's
  vintage and flag anything older than 3 years.
- **Paywalled content**: Many academic papers and industry reports are behind paywalls.
  Extract what is publicly available (title, authors, abstract, key findings from
  summaries) and note that full text requires access.
- **Source credibility**: Note the credibility tier of each source. Government statistics
  and peer-reviewed papers are tier 1; established media and major think tanks are tier 2;
  blogs and self-media are tier 3.

## Error Handling

- **WebSearch returns few results**: Try alternative keywords, synonyms, or broader terms.
  If still insufficient, inform the user and proceed with available results.
- **WebFetch fails on a URL**: Skip that source and try the next one. Do not block the
  entire workflow on a single failed fetch.
- **PDF conversion fails**: Fall back to Word output. Inform the user that PDF requires
  MS Word installed.
- **No results for a category**: If a category yields no usable results, omit that
  category from the brief and note in the summary which categories had no findings.
