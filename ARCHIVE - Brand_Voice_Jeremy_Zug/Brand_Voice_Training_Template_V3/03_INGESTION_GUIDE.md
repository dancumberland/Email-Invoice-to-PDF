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
1.  **Initial Inventory & Tracker Setup (CRITICAL FIRST STEP):**
    a.  Identify ALL client source material files (e.g., from `Raw/`, `Chunks/` if pre-chunked, or a specified client content directory).
    b.  For each identified file, assess if chunking is needed based on the "Chunking Rules" table below. Determine the exact number of chunks for any large files.
    c.  **Crucially, before any processing begins, populate the `Strategic_Context/Process_Tracker.md` with a separate row for *every original file and each of its anticipated chunks*.** Mark these initial entries with a status like 'Pending Ingestion'. This tracker now serves as the complete manifest of all work to be done for the ingestion phase.
2.  **Asset Preparation:**
    a.  Ensure all raw assets listed in the inventory are accessible (e.g., correctly placed in `Raw/`).
    b.  If manual pre-chunking is part of the workflow, ensure these chunks are in `Chunks/` and correctly correspond to the tracker entries. Otherwise, the AI will conceptually chunk files during processing based on the inventory.
3.  **Iterative Ingestion (Processing Loop):**
    a.  Select the next item marked 'Pending Ingestion' from the `Process_Tracker.md`. Process ≤ 3 such chunks per AI interaction session.
    b.  For the selected chunk, run the ***Source Ingestion Prompt*** (detailed below).
    c.  Save the AI-generated summary output to the `Summaries/` directory, using a consistent naming convention (e.g., `originalfilename_partXX_summary.md`).
    d.  **Update the `Process_Tracker.md`**: Change the status of the processed chunk's row to 'Ingestion Complete ✅'. Record the summary filename and any pertinent notes.
4.  **Completion of Ingestion Phase:**
    The ingestion phase is considered complete ONLY when ALL entries initially populated in the `Process_Tracker.md` have been updated to 'Ingestion Complete ✅'.

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
0.  **Perform a Full Inventory Scan & Pre-populate Tracker (Mandatory First Action):**
    *   Before initiating any file processing, comprehensively scan all provided source material locations (e.g., `Raw/`, `Chunks/`, or a specified client content directory).
    *   List every source file. For each, determine the number of chunks required according to the "Chunking Rules."
    *   **Pre-populate the `Strategic_Context/Process_Tracker.md` with entries for ALL identified files and their respective chunks.** This tracker becomes the definitive checklist for the ingestion phase.
1.  **Adhere to Inventory for Processing:** Process files and chunks strictly based on the pre-populated `Process_Tracker.md`.
2.  **Chunk Identification:** If not pre-chunked, identify chunk boundaries for oversized inputs based on the initial inventory assessment.
3.  **Session Limits:** Respect the limit of processing ≤ 3 chunks per interactive session with the AI.
4.  **Safe Input Sizes:** Never pass raw documents exceeding 10k+ tokens (or as per chunking rules) directly to generation prompts; always use the chunked approach.
5.  **Progress Monitoring:** Continuously refer to the `Process_Tracker.md` to determine remaining work and confirm overall completion of the ingestion phase. Ingestion is not complete until every item in the tracker is processed.

---
*Last updated 2025-06-03*
