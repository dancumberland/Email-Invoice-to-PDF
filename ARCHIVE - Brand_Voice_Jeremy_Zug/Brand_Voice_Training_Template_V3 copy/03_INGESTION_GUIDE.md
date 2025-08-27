# Source Ingestion & Chunking Guide (Phase 0)

Large raw assets (transcripts, books, long .txt / .md files) rarely fit inside a single LLM context window.  This guide shows the operator (human or agent) how to turn those assets into compact, reference-ready summaries that *do* fit in context and feed the rest of the Brand Voice V3 system.

---
## 📐 Chunking Rules
| Asset Type | Recommended Chunk | Rationale |
|------------|------------------|-----------|
| Plain text / Markdown | ≤ 3 000 words (≈ 4-5 k tokens) | Keeps each chunk safely under typical 8-16 k windows |
| Audio transcript | ≤ 7 minutes of speech | Similar token footprint; aligns with natural break points |
| Slide decks / PDFs | Split by agenda section or ≤ 30 slides | Keeps logical coherence |

*Always split on sentence / slide boundaries to preserve readability.*  Name each chunk `originalfilename_part01.txt`, `..._part02.md`, etc.

---
## 🔄 Operator Workflow (Phase 0)
1. **Collect Raw Assets** → place in `Raw/`.
2. **Chunk Large Files** per table above → save in `Chunks/`.
3. **Run the *Source Ingestion Prompt*** (below) on ≤ 3 chunks at a time.
4. **Save AI Output** to `Summaries/` as `originalfilename_part01_summary.md`.
5. **Update Process Tracker** (`Strategic_Context/Process_Tracker.md`) by adding/​editing a row.

> ✱ Tip: After ~15 000 processed-token summaries, run the *Voice Calibration Micro-Prompt* (see `02_PROMPTS.md`) to refresh signature patterns.

---
## 📝 Source Ingestion Prompt (copy / paste)
```
SOURCE INGESTION v3
You will read the following *single* chunk of raw brand material and output four sections *exactly* as specified.

1. ## Summary  – A **3-sentence objective** synopsis of this chunk.
2. ## Voice Markers  – Up to **8 bullet points**.  Each bullet = a voice characteristic **plus** a 5-10-word direct quote that exemplifies it.
3. ## Notable Quotes – The **3 strongest verbatim lines** (≤ 40 words each) with line #/​timestamp.
4. ## Tags – 3-8 comma-separated topical tags (e.g. core_values, humor, product_positioning).

After creating the above, append a markdown table row to the Process Tracker with: |source_file|part#|summary_file|Voice Markers Logged| – and leave Notes blank.
```

---
## ✔️ Quality Gate
A chunk is considered *successfully ingested* when:
- A summary file exists in `Summaries/`
- The Process Tracker row is filled with ✅ under *Voice Markers Logged?*

---
## 🤖 Autonomy Guidance  
Agentic tools should:
1. Detect oversized inputs automatically → trigger chunking.
2. Respect 3-chunk limit per session.
3. Never pass raw 10 k+ token documents directly to generation prompts.

---
*Last updated 2025-06-03*
