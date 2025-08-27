# Brand Voice Training Template V4 – START HERE

## 🎯 CORE OBJECTIVE
Create a **single, highly-detailed Master Brand Voice Document (≈30-50 pages)** from raw client materials, then generate equally detailed **Platform-Specific Add-On Guides**.  
This folder contains a minimal file set that works together as a repeatable *system* (similar to V3, simplified for clarity).

| File | Purpose |
| ---- | ------- |
| `00_START_HERE.md` | Quick-start overview + checklist (this file) |
| `01_AI_INSTRUCTIONS.md` | System-level instructions for the LLM |
| `02_PROMPTS.md` | Copy-paste prompt templates (Master + Platform) |
| `03_INGESTION_GUIDE.md` | How to chunk & ingest oversized source docs |
| `04_PROCESS_TRACKER_TEMPLATE.md` | Markdown tracker to monitor progress |

---

## ⚡ QUICK-START CHECKLIST

### Phase 0 – Source Ingestion *(variable)*
1. **Inventory** every raw asset (website copy, transcripts, posts, PDFs, etc.).  
   – List each file in a fresh copy of `04_PROCESS_TRACKER_TEMPLATE.md` under *Client Context Documents*.
2. **Chunk** any file > 10-15 kB using the rules in `03_INGESTION_GUIDE.md`.
3. **Run the *Source Ingestion Prompt*** (see `03_INGESTION_GUIDE.md`) for ≤3 chunks at a time.  
   – Save each summary in a `Summaries/` folder.  
   – Mark that chunk **Ingestion Complete ✅** in the tracker.

> **Ingestion is finished only when *every* row in the tracker is ✅.**

### Phase 1 – Master Document Generation *(≈30-60 min)*
1. Share `01_AI_INSTRUCTIONS.md` with the AI (system message).  
2. Copy the **Master Document Prompt** from `02_PROMPTS.md` into the chat.  
3. Let the AI iterate through all summaries, populating the 17-section structure.
4. Review, request edits, and save as `Generated_Brand_Voice_Master.md` in `Master_Guide/`.

### Phase 2 – Platform Guides *(≈30-60 min each)*
1. For each desired platform, copy the corresponding **Platform Guide Prompt** from `02_PROMPTS.md` and update placeholders.  
2. Review output, request refinements, save in `Platform_Guides/`.

### Phase 3 – Delivery & QC
1. Run through any internal quality checklist.  
2. Package the Master Guide + Platform Guides for client delivery.

---

## 🛠  FOLDER CONVENTION (RECOMMENDED)
```
Brand_Voice_[ClientName]/
├─ Master_Guide/
│   └─ Generated_Brand_Voice_Master.md
├─ Platform_Guides/
│   └─ LinkedIn_Guide.md (etc.)
├─ Raw/               # original source assets
├─ Chunks/            # optional pre-split text chunks
├─ Summaries/         # AI-generated chunk summaries
└─ Strategic_Context/ # (optional) North Star, ICP, Offers, etc.
```

---

## 🔑 SUCCESS CRITERIA
• Master Guide covers **ALL 17 sections** with required examples, counter-examples, & exercises.  
• Platform Guides adapt voice nuances with deep, annotated examples.  
• Any competent AI can load the guides and instantly mimic the client’s voice.

> *Last updated: 2025-06-04 | Template V4.0*
