---
name: eunseo-resume-tailor
description: >
  Tailors Eunseo Jeong's base resume to a target Job Description and produces
  an ATS-optimized Word (.docx) file. Trigger whenever the user provides a JD
  (plain text or URL) and asks to tailor, customize, or rewrite the resume for
  a specific role. The agent must never fabricate skills, titles, metrics, or
  achievements not present in the source resume.
---

# Eunseo Resume Tailoring Skill

## Purpose

Generate a job-tailored `.docx` resume from:
1. **Source resume** — single source of truth (Eunseo's base resume, embedded below)
2. **Job description** — plain text or URL

Output must be:
- ATS-friendly (plain text, standard headings, no tables/graphics)
- Recruiter-scannable in 6–8 seconds
- 100% traceable to the source resume — no fabrication

---

## Candidate Context (for reference only — do not output this section)

**Eunseo Jeong** is a Product Manager with:
- Samsung Electronics: PM on world-first Smart Fridge advertising surface (U.S. market), Sep 2024–Present
- Naver Cloud: PM Intern, AI Cloud Platform (LLM evaluation, competitive analysis), Jun 2024
- Mathpresso: UX Research Intern, Qanda EdTech app (VOC analysis, usability testing), Aug–Dec 2021
- Act Startup Consulting Group: Go-to-market strategy, U.S. market entry, Mar–Dec 2022
- Seoul National University: BA Communication / Information Science & Culture, GPA 3.9
- University of Washington: Study Abroad, 2023
- Published: "Code-to-Design: Bridging the Gap with Design Tokens" — Smashing Magazine (2021)

Key strengths: advertising product operations, campaign performance analysis, SQL-based data analysis,
AI product strategy, LLM evaluation, UX research, VOC analysis, cross-functional stakeholder alignment,
go-to-market support, Korean native + English professional proficiency.

---

## Step-by-Step Workflow

Follow these steps **in order** every time.

---

### Step 1 — Parse the Job Description

If a URL is provided, fetch the page content first.

Extract and record:
| Element | Purpose |
|---|---|
| Role title | Match seniority and function |
| Core responsibilities | Map to source resume bullets |
| Required skills & tools | Direct ATS keyword targets |
| Preferred / nice-to-have | Secondary relevance signals |
| Industry / domain signals | Ads, AI, Data, Growth, UX, Marketing, etc. |
| Repeated keywords | Highest ATS priority |

---

### Step 2 — Relevance Mapping

Compare the JD to the source resume. For each JD requirement, classify:

- ✅ **Covered** — source resume has clear, direct evidence → include and emphasize
- ⚠️ **Partial** — adjacent experience exists; reframe carefully using only source language
- ❌ **Not covered** — no basis in source resume → do NOT include; note as gap if asked

Document the mapping internally before rewriting.

---

### Step 3 — Rewrite Resume Content

**Rules:**
- Maintain the **exact same section order** as the source resume:
  Summary → Experience → Project Experience → Skills & Tools → Education → Recognition
- Reorder bullets **within** each role to lead with the most JD-relevant content
- Use JD keywords **only where the source resume supports the claim**
- Remove or de-emphasize bullets with low relevance to this JD
- Keep all numbers and metrics exactly as they appear in the source (e.g., 0.5 points, 30%, 300+, 36 beta customers)
- Rewrite Summary to position Eunseo for this specific role and industry
- Translate internal terminology into externally understandable language
- Remove internal acronyms, log names, and tool-specific jargon that recruiters won't recognize

**Bullet structure:**
```
Action Verb + What was done + Context/Method + Result/Impact
```

Example:
- Analyzed U.S. advertising campaign performance across Smart Fridge screen placements, identifying usage patterns by time of day and content type to inform inventory and scheduling strategy.

**Forbidden:**
- Inventing any skill, tool, metric, job title, or achievement
- Inferring experience not explicitly stated in the source
- Changing job titles, company names, or date ranges
- Overstating seniority or ownership beyond what the source supports
- Adding tools (e.g., Tableau, dbt, Python libraries) not mentioned in the source

**Length:**
- Eunseo has ~2 years of full-time experience → target **1 page** where possible
- Allow 2 pages only if the JD demands extensive coverage of multiple domains

---

### Step 4 — Generate the .docx

Use the `docx` npm package. Install with: `npm install -g docx`

#### Key constants
```javascript
const FONT = "Calibri";
const NAME_SIZE = 36;      // 18pt
const CONTACT_SIZE = 18;   // 9pt
const BODY_SIZE = 20;      // 10pt
const SECTION_SIZE = 20;   // 10pt, allCaps, bold
```

#### Page setup (US Letter, 0.75" margins)
```javascript
properties: {
  page: {
    size: { width: 12240, height: 15840 },
    margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
  }
}
```

#### Numbering config (add to Document root)
```javascript
numbering: {
  config: [{
    reference: "bullets",
    levels: [{
      level: 0,
      format: LevelFormat.BULLET,
      text: "•",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 360, hanging: 360 } } }
    }]
  }]
}
```

#### Section header with bottom border
```javascript
new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000", space: 4 } },
  spacing: { before: 180, after: 60 },
  children: [new TextRun({
    text: "EXPERIENCE",
    bold: true, size: SECTION_SIZE, font: FONT, allCaps: true
  })]
})
```

#### Job title row with right-aligned dates
```javascript
new Paragraph({
  tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
  spacing: { before: 80, after: 30 },
  children: [
    new TextRun({ text: "Product Manager – Ads & Service", bold: true, size: BODY_SIZE, font: FONT }),
    new TextRun({ text: " | Samsung Electronics | Suwon, South Korea", size: BODY_SIZE, font: FONT, italics: true }),
    new TextRun({ text: "\tSep 2024 – Present", size: BODY_SIZE, font: FONT }),
  ]
})
```

#### Italic role descriptor line (appears below job title)
```javascript
new Paragraph({
  spacing: { before: 0, after: 30 },
  children: [new TextRun({
    text: "Building and analyzing a world-first advertising surface on Smart Fridge screens for the U.S. market.",
    size: BODY_SIZE, font: FONT, italics: true
  })]
})
```

#### Bullet paragraph
```javascript
new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { before: 30, after: 30 },
  children: [new TextRun({ text: "Bullet text here.", size: BODY_SIZE, font: FONT })]
})
```

#### Validation
```bash
python /path/to/docx/scripts/office/validate.py output.docx
```

---

### Step 5 — Quality Check

Before delivering, verify:

- [ ] Every claim traces to the source resume
- [ ] No skills, tools, or metrics were added beyond the source
- [ ] Section order matches: Summary → Experience → Project → Skills → Education → Recognition
- [ ] All bullets follow Action + What + Context + Impact structure
- [ ] JD keywords appear naturally, not stuffed
- [ ] Numbers are exact (0.5 points, 30%, 300+, 36 customers, 500+ hours)
- [ ] Internal jargon translated to recruiter-readable language
- [ ] File validates without errors
- [ ] Length is 1 page (or justified 2-page if needed)

---

## Output

1. **Tailored .docx file** saved to outputs folder
2. **Brief tailoring rationale** (unless user requests "resume only"):
   - JD requirements emphasized and how they map to source
   - Any JD requirements not covered (no source basis)
3. **File naming:**
```
eunseo-jeong-[company]-[role-slug].docx
# Examples:
eunseo-jeong-google-pm-ads.docx
eunseo-jeong-kakao-product-manager.docx
```

---

## Guardrails

| ❌ Never | ✅ Always |
|---|---|
| Invent skills, tools, or metrics | Use verbatim from source resume |
| Change job titles or companies | Keep exact titles from source |
| Add uncovered JD requirements | Mark as gap; don't include |
| Use tables for layout | Use tab stops and paragraph borders |
| Use `\n` inside Paragraph | Create separate Paragraph objects |
| Use unicode bullet characters manually | Use `LevelFormat.BULLET` with numbering config |
| Overstate seniority or ownership | Reflect actual role scope from source |
| Include internal Samsung/Naver acronyms | Translate to external-friendly language |

---

## Source Resume Reference

Use this embedded content as the single source of truth when the user does not upload a fresh resume.

```
EUNSEO JEONG
Seoul, South Korea | 010-3449-5029 | dpffltm5029@email.com | linkedin.com/in/eunseoj/

SUMMARY
Data-driven Product Manager with experience in advertising products, AI product strategy,
product marketing, and UX research. Currently working on a world-first smart refrigerator
advertising surface for the U.S. market, leveraging device-level product data and campaign
performance analysis to understand user behavior, support business decision-making, and
improve advertising inventory strategy. Skilled at translating ambiguous product problems
into actionable insights across product, data, engineering, ad operations, and regional
business teams.

EXPERIENCE

Product Manager – Ads & Service | Samsung Electronics | Suwon, South Korea | Sep 2024 – Present
Building and analyzing a world-first advertising surface on Smart Fridge screens for the
U.S. market, creating a new media opportunity beyond traditional mobile, web, and CTV.

- Analyzed performance of U.S.-targeted advertising campaigns across Smart Fridge screen
  placements, including impressions, clicks, dismiss actions, CTR, dismiss rate, unique
  reach, and device-level engagement trends.
- Leveraged internal appliance event data and SQL-based analysis to understand how users
  engage with a new screen-based advertising surface, translating behavioral patterns into
  business insights for advertising strategy and product decision-making.
- Supported beta campaign performance reviews across multiple brand categories, connecting
  DSP campaign data with internal product usage data to explain reach, engagement, user
  response, and placement-level performance.
- Identified key usage patterns by placement, time of day, campaign period, and content
  engagement level, helping inform ad scheduling, inventory positioning, prime-time
  strategy, and frequency cap discussions.
- Built AI-assisted reporting and dashboard workflows using SQL, Superset, Excel, VBA,
  Claude, ChatGPT, and Gemini, reducing manual analysis effort and improving consistency
  across recurring campaign performance reviews.
- Partnered with product, engineering, data, ad operations, legal, and regional business
  stakeholders to align on logging requirements, metric definitions, campaign analysis
  needs, and go-to-market support for a new advertising product.

Product Management Intern – AI Cloud Platform | Naver Cloud | Jeong Ja, South Korea | Jun 2024
Supported AI product strategy and evaluation for Naver's cloud and large language model products.

- Developed a roadmap for new AI product lines and positioning strategies through market
  and competitive analysis, including comparisons with Microsoft, AWS, and Google.
- Established evaluation metrics for Naver's large language model, HyperClova X, by
  analyzing diverse LLM benchmarks to support more precise model performance evaluation.
- Synthesized AI market trends, competitor capabilities, and internal product strengths
  into strategic recommendations for AI product development and positioning.

UX Research Intern | Mathpresso | Seoul, South Korea | Aug 2021 – Dec 2021
Conducted user research and VOC analysis for Qanda, a large-scale EdTech app.

- Improved Qanda's App Store rating by over 0.5 points by analyzing 300+ Voice of Customer
  feedback items and prioritizing usability improvements.
- Increased conversion rates by 30% by identifying an emotional barrier in the "Request
  for Parents" subscription onboarding flow through 7 in-depth interviews and usability tests.
- Collaborated with data scientists, designers, and developers to create customer journey
  maps, expand use cases, and support product improvements for the learning experience.

PROJECT EXPERIENCE

Act Startup Consulting Group | Student Associate Consultant | Seoul, South Korea | Mar 2022 – Dec 2022
Developed go-to-market and marketing strategies for early-stage business expansion projects.

- Contributed to go-to-market strategy development for the U.S. expansion of a Korean
  jewelry company by combining market analysis, user research, and strategy consulting frameworks.
- Developed a seasonal multi-channel marketing strategy focused on Instagram, using the
  AARRR framework and analysis of 10+ U.S. competitor accounts.
- Acquired 36 beta customers through a referral coupon event, leveraging insights from 20+
  in-depth interviews with U.S. target customers to understand jewelry consumption behavior.

SKILLS & TOOLS

Product: Product Strategy, Product Roadmapping, Product Requirements, Stakeholder Communication,
  Go-to-Market Support, Advertising Product Operations, User Behavior Analysis

Data & Analytics: SQL, Python, BigQuery, Superset, Excel, VBA, Product Usage Data Analysis,
  Campaign Performance Analysis, KPI Analysis, Funnel Analysis, User Segmentation, Data Storytelling

Advertising & Business: Ad Inventory Strategy, Smart Screen Advertising, Campaign Analysis,
  CTR Analysis, Dismiss Rate Analysis, Frequency Cap Analysis, DSP Data Analysis, U.S. Market Research

Research: User Interviews, Usability Testing, Concept Testing, Focus Group Discussions,
  VOC Analysis, Journey Mapping, Persona Development, A/B Testing

Tools: Figma, Miro, Jira, Notion, PowerPoint, Claude, ChatGPT, Gemini, Cursor AI

Languages: Korean (Native), English (Professional Working Proficiency)

EDUCATION

Seoul National University | Seoul, South Korea
Bachelor of Communication / Information Science and Culture | Mar 2019 – Feb 2024 | GPA: 3.9

University of Washington | Seattle, WA
Study Abroad | Mar 2023 – Jul 2023

RECOGNITION

Published: "Code-to-Design: Bridging the Gap with Design Tokens" — Smashing Magazine (2021).
```
