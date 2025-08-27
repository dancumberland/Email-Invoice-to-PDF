# SOP: Generating Brand Voice Guides (V3 Template System)
Version 1.0 | 2025-06-03

This document outlines the standard operating procedure and checklist for generating a concise Master Brand Voice Guide and modular Platform-Specific Add-On Guides using the V3 template system. This process involves collaboration between a human operator (Content Lead/PM) and an AI assistant.

**Key Reference Documents:**
*   `00_AI_INSTRUCTIONS.md` (Version 2.0 - Concise & Modular Focus)
*   `prompt_skeletons_for_guide_generation.md`
*   `template_creation_process.md` (for overall workflow and QC points)

---
## Phase 1: Preparation & Setup

*   [ ] **1.1. Identify Pilot Client:** Confirm the client for whom the brand voice guides will be generated.
    *   Client Name: _________________________
*   [ ] **1.2. Gather Client Source Materials:**
    *   [ ] Collect all relevant existing brand documentation, content examples, and raw materials from the client. This may include previous brand guides, website copy, blog posts, social media content, marketing materials, mission statements, value propositions, audience research, etc.
    *   *Goal: Obtain the equivalent information that would have populated the OLD, more granular template files (e.g., `01_executive_summary.md`, `02_brand_foundation.md`, `05_linguistic_patterns.md`, `04_content_generation_guidelines.md` for platform specifics, etc.).*
*   [ ] **1.3. Organize Source Materials:**
    *   [ ] Convert materials to AI-processable formats (e.g., plain text, markdown) if necessary.
    *   [ ] Group related source files or content snippets logically to align with the input requirements of the prompt skeletons (e.g., gather all 'core identity' related docs together).
*   [ ] **1.4. Create Client Folder Structure:**
    *   [ ] Create the main client directory: `/Brand_Voice_[ClientName]/`
    *   [ ] Inside, create `Master_Guide/` subdirectory.
    *   [ ] Inside, create `Platform_Guides/` subdirectory.

---
## Phase 2: Master Brand Voice Guide Generation (`Generated_Brand_Voice_Master.md`)

This is an iterative process. For each section, you will provide the AI with the relevant prompt and source materials, review the output, and request revisions as needed.

*   [ ] **2.1. Generate Section: `00_How_To_Use_This_Guide_For_AI`**
    *   [ ] Provide AI with the prompt from `prompt_skeletons_for_guide_generation.md`.
    *   [ ] Review AI output for clarity, conciseness, and accuracy.
    *   [ ] Iterate with AI until satisfactory. Append to `Generated_Brand_Voice_Master.md`.
*   [ ] **2.2. Generate Section: `01_Core_Brand_Identity`**
    *   [ ] Provide AI with the prompt and relevant client source materials (e.g., old `01_executive_summary.md`, `02_brand_foundation.md`, `03_brand_assets_and_history.md` equivalents).
    *   [ ] Emphasize adherence to `00_AI_INSTRUCTIONS.md` (conciseness, focus on voice-informing elements).
    *   [ ] Review AI output. Iterate until satisfactory. Append to `Generated_Brand_Voice_Master.md`.
*   [ ] **2.3. Generate Section: `02_Voice_Tone_And_Style`**
    *   [ ] Provide AI with the prompt and relevant client source materials (e.g., old `05_linguistic_patterns.md`, `08_brand_language_parameters.md` equivalents).
    *   [ ] Review AI output. Iterate until satisfactory. Append to `Generated_Brand_Voice_Master.md`.
*   [ ] **2.4. Generate Section: `03_Content_Strategy_And_Audience`**
    *   [ ] Provide AI with the prompt and relevant client source materials (e.g., old `04_content_generation_guidelines.md`, `06_audience_engagement.md` equivalents).
    *   [ ] Review AI output. Iterate until satisfactory. Append to `Generated_Brand_Voice_Master.md`.
*   [ ] **2.5. Generate Section: `04_Key_Exemplars_And_Storytelling`**
    *   [ ] Provide AI with the prompt and relevant client source materials (e.g., old `07_advanced_storytelling.md` and BEST actual content examples from client).
    *   [ ] Emphasize selection of 2-4 *short, highly representative* examples with brief annotations.
    *   [ ] Review AI output. Iterate until satisfactory. Append to `Generated_Brand_Voice_Master.md`.
*   [ ] **2.6. Holistic Review of Master Guide:**
    *   [ ] Read through the complete `Generated_Brand_Voice_Master.md`.
    *   [ ] Check for overall consistency, flow, and clarity.
    *   [ ] Verify adherence to total length target (10-20 pages).
    *   [ ] Make any final human edits or request AI revisions for overall polish.

---
## Phase 3: Platform-Specific Add-On Guide Generation

For each platform required by the client (e.g., LinkedIn, Newsletter, Website Blog, Instagram):

*   [ ] **3.1. Identify Target Platform:** _________________________
*   [ ] **3.2. Generate `[PlatformName]_Guide.md`:**
    *   [ ] Provide AI with the 'General Prompt Structure for ANY Platform-Specific Add-On Guide' from `prompt_skeletons_for_guide_generation.md`.
    *   [ ] Customize the prompt with the specific `[Client Name]` and `[Platform Name]`.
    *   [ ] Provide AI with relevant client source materials containing platform-specific strategies, content examples, or nuances (e.g., sections from old `04_content_generation_guidelines.md`, actual platform posts).
    *   [ ] **Crucially instruct AI:** This guide *supplements* the Master Guide, should *not repeat* its content, and must focus only on *deltas and specific applications* for the platform. Target 3-5 pages.
    *   [ ] Review AI output for conciseness, platform-specificity, and accurate referencing of the Master Guide (implicitly or explicitly).
    *   [ ] Iterate with AI until satisfactory.
    *   [ ] Save the final guide as `[PlatformName]_Guide.md` in the `/Brand_Voice_[ClientName]/Platform_Guides/` directory.
*   [ ] **3.3. Repeat for other platforms as needed.**
    *   Platform 2: _________________________ [ ] Generated [ ] Reviewed
    *   Platform 3: _________________________ [ ] Generated [ ] Reviewed
    *   (Add more as needed)

---
## Phase 4: Quality Control (QC)

Refer to the QC checklist in `template_creation_process.md` (Section 6) for detailed criteria.

*   [ ] **4.1. Master Guide QC:**
    *   [ ] Length within target (10-20 pages)?
    *   [ ] Core voice characteristics clearly defined and present?
    *   [ ] Examples are real (or highly representative) and illustrative?
    *   [ ] No redundant sections; information is concise and high-signal?
*   [ ] **4.2. Platform Add-On Guides QC (for each guide):**
    *   [ ] Length within target (3-5 pages)?
    *   [ ] Clearly references the Master Guide (implicitly or explicitly) and avoids duplication?
    *   [ ] Focuses on platform-specific deltas and nuances?
    *   [ ] Examples are specific to the platform and demonstrate voice adaptation?
*   [ ] **4.3. Overall System QC:**
    *   [ ] All files correctly named and in the proper folder structure?
    *   [ ] Versioning clear (e.g., v1.0 for initial client delivery)?

---
## Phase 5: Finalization & Delivery

*   [ ] **5.1. Package Deliverables:** Prepare the `Generated_Brand_Voice_Master.md` and all `[PlatformName]_Guide.md` files for the client.
*   [ ] **5.2. Client Review & Sign-off (if applicable):** Coordinate client review and incorporate any final minor feedback.
    *   Client Sign-off Date: _______________
*   [ ] **5.3. Internal Archiving:** Archive the final approved versions of all generated guides and relevant source materials internally.

---
**SOP Review & Updates:** This SOP should be reviewed periodically and updated as the V3 template system and processes evolve.
