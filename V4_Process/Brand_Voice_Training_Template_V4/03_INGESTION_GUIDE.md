# 03_INGESTION_GUIDE.md – Chunking & Ingestion (Template V4)

## ⚙️ WHEN TO CHUNK
Chunk any file that is:
* **> 15 kB** (plain text) or roughly **>5,000 words**
* A long transcript, book chapter, or dense article

## 🔪 CHUNKING RULES
1. Aim for **3-5 kB (~1,000-1,500 words)** per chunk.
2. Preserve logical breaks (paragraphs, speaker changes, headings).
3. Name chunks sequentially: `OriginalFile__Chunk01.txt`, `...Chunk02.txt`, etc.
4. Store in `Chunks/` and record each chunk in the Process Tracker.

## 🚀 CHUNK APPLICATION PROMPT (no summaries)
Copy this into the chat *for each chunk* (max 3 per session):

```
SYSTEM: You are an expert brand voice analyst.

USER: Analyze the following chunk from [FileName]. For every relevant insight, directly update the current draft of the **Master Brand Voice Document** (17-section template). Specifically:
• Identify verbatim quotes, vocabulary, tone signals, stylistic patterns.
• Decide which section/sub-section the insight belongs to and integrate it there.
• When adding examples or counter-examples, enclose verbatim text in blockquotes.
• DO NOT output a standalone summary. Instead, return ONLY the updated portions of the Master Document (or the entire doc if easier) in a single Markdown code block.
• Mark this chunk as **Ingestion Complete ✅** in the Process Tracker.
• After reading, set Status to **🟡 Read** in the Process Tracker.
• Once integrated into the Master Guide, change the Status to **✅ Applied**.

[PASTE CHUNK HERE]
```

**Next Steps:** After the AI integrates the chunk into the Master Document, copy the updated content into `Master_Guide/Generated_Brand_Voice_Master.md` (overwriting previous draft). Then mark the chunk **✅ Applied** in the tracker.

*Last updated 2025-06-04 · Template V4.2*
