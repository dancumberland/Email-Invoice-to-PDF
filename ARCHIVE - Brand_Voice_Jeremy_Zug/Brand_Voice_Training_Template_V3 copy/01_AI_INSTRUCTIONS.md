# AI Implementation Instructions for Brand Voice Guide Generation
Version 2.0 (Concise & Modular Focus)

## CORE OBJECTIVE: Generate a Concise, High-Signal Brand Voice Guide

Your primary goal is to synthesize the provided brand materials into a **brief yet potent** master brand voice guide (target: 10-20 pages total) and several **lightweight, standalone platform-specific add-on guides** (target: 3-5 pages each).

**Prioritize ruthlessly:** Extract the absolute essence of the brand's voice. Focus on clarity, impact, and actionable examples. Avoid redundancy and excessive detail. The final documents must be easily digestible by another AI within a limited context window.

---

## 1. Master Brand Voice Guide: Structure & Content

Create a single, consolidated `Generated_Brand_Voice_Master.md`. This document should synthesize information into the following core sections. **Do NOT create separate files for these sections initially; they are logical divisions within the single master document.**

   ```markdown
   # [Client Name] - Master Brand Voice Guide
   
   ## 00_How_To_Use_This_Guide_For_AI
   - Brief (1-2 paragraphs) instructions for the consuming AI.
   - Example: "You are [Brand Name]'s AI voice. Your primary goal is to generate content that is [key characteristic 1] while being [key characteristic 2]. Refer to the examples and guidelines below. Prioritize authenticity and direct application of these patterns."
   
   ## 01_Core_Brand_Identity
   - **Source Synthesis From:** `01_executive_summary.md`, `02_brand_foundation.md`, relevant parts of `03_brand_assets_and_history.md`.
   - **Content Focus:**
     - Mission (1-2 sentences)
     - Core Values (3-5 bullet points)
     - Brand Essence/Personality (3-5 descriptive adjectives with 1 brief explanatory sentence each)
     - Unique Selling Proposition (1 concise statement)
     - Brief Brand Story/Origin (if it *directly and significantly* informs the voice, 1-2 short paragraphs max).
   
   ## 02_Voice_Tone_And_Style
   - **Source Synthesis From:** `05_linguistic_patterns.md` (condensed), `08_brand_language_parameters.md` (key parts), voice elements from `02_brand_foundation.md`.
   - **Content Focus:**
     - Overall Voice Characteristics (e.g., Warm, Authoritative, Witty - choose 3-5, provide 1-2 *brief, illustrative examples* for each from source material).
     - Tone Spectrum (e.g., Conversational yet Professional; list 2-3 key spectrums with a short explanation).
     - Preferred Language/Vocabulary (List 5-10 *key* preferred words/phrases with 1 example each. Focus on unique or highly characteristic terms).
     - Discouraged Language/Vocabulary (List 5-10 *key* discouraged words/phrases with 1 example each or a brief reason).
     - Signature Phrases/Taglines (If any, list them).
     - Sentence & Paragraph Style (Brief notes on preferred length, complexity, rhythm if distinctly characteristic. 1-2 examples).
   
   ## 03_Content_Strategy_And_Audience
   - **Source Synthesis From:** Key principles from `04_content_generation_guidelines.md` (general principles, not deep platform specifics), `06_audience_engagement.md` (audience summary, core engagement principles).
   - **Content Focus:**
     - Core Content Themes/Pillars (List 3-5 primary themes).
     - Primary Audience Snapshot (Brief description: who they are, key needs/motivations relevant to voice. 2-3 sentences per key segment if distinct).
     - General Approach to Content (e.g., "Educate and empower through storytelling," "Provide actionable insights with a direct, no-nonsense approach.").
     - Key Calls to Action (General CTAs, not platform-specific yet. List 2-3 common ones).
   
   ## 04_Key_Exemplars_And_Storytelling
   - **Source Synthesis From:** Best parts of `07_advanced_storytelling.md` (if a few core story archetypes are vital), and *most importantly, direct examples from client's best content*.
   - **Content Focus:**
     - Showcase, Don't Just Describe: Provide 2-4 *short, highly representative* examples (e.g., a paragraph, a social media post snippet) of the brand's content in action that clearly demonstrate the voice.
     - Minimal annotations on *why* the exemplar is good (e.g., "Note the use of [key characteristic] and [preferred phrase]").
     - Core Story Archetypes (If central to the brand, briefly describe 1-2 archetypes with a short example of how they manifest).
   ```

---

## 2. Platform-Specific Add-On Guides: Structure & Content

After the Master Guide is drafted, you will be prompted to create **separate, lightweight `.md` files** for each specified platform (e.g., `LinkedIn_Guide.md`, `Newsletter_Guide.md`). These should be 3-5 pages MAX.

**Key Principles for Platform Guides:**
*   **Reference, Don't Repeat:** These guides should *briefly reference* core principles from the Master Guide and only detail the *deltas* or specific applications for that platform.
*   **Focus on Actionable Differences:** What *specifically* changes or needs emphasis for this platform?

   **Template for Platform-Specific Guides:**
   ```markdown
   # [Platform Name] - Brand Voice Add-On
   
   **Reference:** This guide supplements the [Client Name] Master Brand Voice Guide. Refer to the Master Guide for core identity, voice, and style principles.
   
   ## 1. Purpose & Goals on [Platform Name]
   - What are the primary objectives for this brand on this platform? (1-2 sentences)
   
   ## 2. Content Structure & Format on [Platform Name]
   - Specific post/content structures (e.g., LinkedIn post: Hook, Value, CTA).
   - Formatting nuances (e.g., use of emojis, hashtags, video length).
   
   ## 3. Tone & Voice Nuances for [Platform Name]
   - Any specific shifts in tone? (e.g., "Slightly more informal on Instagram DMs").
   - Platform-specific Do's and Don'ts (that aren't covered in Master).
   
   ## 4. Key Exemplar(s) for [Platform Name]
   - 1-2 *short, specific examples* of successful content for this brand *on this platform*.
   
   ## 5. Calls to Action (CTAs) & Engagement for [Platform Name]
   - Platform-specific CTAs.
   - Common engagement strategies for this platform.
   ```

---

## 3. AI Content Generation Rules (NEW)

*   **Overall Length (Master Guide):** Aim for **10-20 pages total**. Be concise.
*   **Overall Length (Platform Guides):** Aim for **3-5 pages total EACH**. Be very concise.
*   **Synthesis is Key:** Do not simply copy-paste. Analyze, extract the essence, and rephrase concisely.
*   **Examples:** Use **1-3 highly illustrative, brief examples** per key point. Prefer direct quotes or short snippets from source material over lengthy new creations.
*   **Show, Don't Just Tell:** Where possible, use an example to make a point rather than a long explanation.
*   **Avoid Redundancy:** If a point is made in one section, do not repeat it extensively elsewhere. Cross-reference if necessary using section titles.
*   **Clarity and Actionability:** Every piece of information should be clear, easy to understand, and actionable for the consuming AI.

---

## 4. Process & Quality

1.  **Initial Analysis:** Review all provided brand materials. Identify core voice patterns, unique brand elements, and standout examples.
2.  **Draft Master Guide:** Generate the `Generated_Brand_Voice_Master.md` following the structure and content guidelines above. Focus on synthesis and brevity.
3.  **Draft Platform Guides:** Once the Master Guide is approved, generate the platform-specific add-on guides.
4.  **Self-Correction/Refinement:** Review your output. Is it concise? Is it clear? Are there redundancies? Does it meet the length targets? Refine as needed.

---

## FINAL INSTRUCTION: Prioritize Brevity and Impact

Your main goal is to create training documents that are **efficient and effective** for another AI to use. Every word should count. When in doubt, choose the more concise way to express an idea. Extract the signal, eliminate the noise. The success of this system relies on your ability to synthesize complex brand information into a powerful, streamlined guide.
