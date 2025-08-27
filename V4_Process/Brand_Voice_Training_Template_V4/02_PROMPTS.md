# 02_PROMPTS.md – Prompt Library (Template V4)

## 1️⃣ MASTER DOCUMENT PROMPT
```
SYSTEM MESSAGE (paste `01_AI_INSTRUCTIONS.md` here)

USER PROMPT:
You are an expert brand strategist. Using ALL summaries in `Summaries/`, create a **30-50 page Master Brand Voice Document** for [THOUGHT LEADER] following the 17-section structure below.

### 17-Section Outline (include each in full):
1. Executive Summary & Document Purpose
2. Brand Voice Foundation
3. Linguistic Analysis
4. Tone Mapping
5. Content Type Specifications (general)
6. Technical Style Guide
7. Audience Understanding
8. Messaging Framework
9. Content Strategy Integration
10. Storytelling Elements
11. Brand Language Parameters
12. Social Media Guidelines (general)
13. Educational Content Development
14. Trust Building Elements
15. Sales & Marketing Integration
16. Content Quality Control
17. Practical Applications & Scenarios

**For EVERY subsection:** Definition & Purpose | 2-3 Examples | 1-2 Counter-Examples | Implementation | Success Criteria | Pitfalls | Exercises.

Return the entire document in one Markdown code block. Conclude with: *We got it all Captain, Sir!*.
```

---

## 2️⃣ PLATFORM GUIDE PROMPT
```
SYSTEM MESSAGE (paste `01_AI_INSTRUCTIONS.md` here)

USER PROMPT:
Using the approved Master Brand Voice Document, create a **[Platform Name] Add-On Voice Guide** (≈15-30 pages) for [THOUGHT LEADER]. Follow the Platform Guide structure (sections 1-8) in `01_AI_INSTRUCTIONS.md`.

Return in one Markdown code block. End with: *We got it all Captain, Sir!*.
```

---

## 3️⃣ SOURCE INGESTION PROMPT  (reference only – see `03_INGESTION_GUIDE.md`)
```
SYSTEM: You are a summarizer.
USER: Summarize the following chunk from [FileName]. Focus on voice cues, vocabulary, tone, stylistic patterns, audience insights, and any notable quotes. Keep ≤300 words.

[PASTE CHUNK HERE]
```

*Last updated 2025-06-04 · Template V4.0*
