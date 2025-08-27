# Template Creation Process
Version 0.1 | 2025-06-03

This living document tracks the end-to-end plan for rebuilding the Brand Voice Training template so it produces concise, high-signal guides plus modular platform add-ons. Update this checklist as tasks are completed or refined.

---
## 1. Objectives
1. Create a **concise (≈10–20 pages)** master Brand Voice guide that an LLM can load in a single context window.
2. Provide **separate, lightweight (≈3–5 pages)** platform-specific add-on files (LinkedIn, Newsletter, Website, etc.).
3. Maintain clarity, exemplary patterns, and implementation guidance without unnecessary verbosity.
4. Build an internal process that is repeatable for any client and easy to quality-control.

---
## 2. High-Level Workflow
| Step | Owner | Key Outputs |
| ----- | ------ | ------------- |
| 1. Source Material Collection | PM | Links/files (podcasts, blog posts, social, etc.) |
| 2. Pre-Processing | AI + PM | Transcripts & text corpus |
| 3. **Revise `00_AI_INSTRUCTIONS.md`** | Content Lead | Slimmed instructions emphasising synthesis & brevity |
| 4. **Consolidate Template Structure** | Content Lead | Mapped sections → `01_Core_Brand_Identity.md`, etc. |
| 5. Generate Draft Master Guide | AI | `Generated_Brand_Voice_Master.md` (v2) |
| 6. Human Review & Redline | SME | Inline edits / comments |
| 7. Iterate & Finalize Master Guide | AI | Approved master doc |
| 8. Create Platform Add-ons | AI | e.g. `LinkedIn_Guide.md`, `Newsletter_Guide.md` |
| 9. Quality Control Pass | QC Lead | QC checklist completed |
| 10. Delivery & Archive | PM | Client package & internal archive |

---
## 3. File & Folder Standards
* **Master Guide Folder:** `/Brand_Voice_[Client]/Master_Guide/`
* **Platform Add-ons:** `/Brand_Voice_[Client]/Platform_Guides/{platform}.md`
* **Naming:** snake_case, no spaces (e.g., `core_brand_identity.md`).
* **Versioning:** Semantic (v1.0 = first client release). Increment minor for content edits, patch for typos.

---
## 4. Content Length Targets
| Section | Max Length |
| ------ | ----------- |
| Master Guide Total | ≈10-20 pages (≈4-6k tokens) |
| Each Platform Guide | ≈3-5 pages (≈1-2k tokens) |
| Examples per Point | 1-3 illustrative excerpts |

---
## 5. Key Deliverables Definition
1. **Core_Brand_Identity.md** – Mission, values, voice characteristics (bullet + 2-3 examples).
2. **Voice_Tone_And_Style.md** – Preferred/avoided language, signature phrases, tone spectrum.
3. **Content_Strategy_And_Audience.md** – Content pillars, primary audience snapshot, engagement principles.
4. **Key_Exemplars_And_Storytelling.md** – 2-4 annotated “gold standard” excerpts.
5. **Platform Guides** – LinkedIn, Newsletter, Blog, etc.  Structure template:
   ```md
   # [Platform] Voice Add-On
   ## Purpose & Goals
   ## Post / Content Structure
   ## Tone Nuances & Dos/Don’ts
   ## Example Post(s)
   ## CTAs & Engagement
   ```

---
## 6. Quality Control Checklist (abbrev.)
- [ ] Length within targets
- [ ] Core voice characteristics present
- [ ] Examples are real & representative
- [ ] No redundant sections
- [ ] Platform guides reference master doc (no duplicate content)

Full QC details remain in `09_quality_control_system.md`.

---
## 7. Open Questions / Risks
* **Token Budget:** Monitor actual token count of generated docs.
* **Example Sourcing:** Ensure usage rights for direct quotes.
* **Stakeholder Sign-Off:** Define who approves the final master guide.

---
## 8. Next Actions (as of 2025-06-03)
- [x] Revise `00_AI_INSTRUCTIONS.md` to reflect new synthesis focus. (Completed)
- [x] Draft mapping table: old template files → new consolidated sections. (Effectively completed via `00_AI_INSTRUCTIONS.md` and prompt skeletons)
- [x] Prepare prompt skeletons for each new section. (Completed: `prompt_skeletons_for_guide_generation.md`)
- [ ] Identify first client to pilot the new flow.

> *This document is maintained by the Template Ops team.  Update consistently to reflect progress and new insights.*
