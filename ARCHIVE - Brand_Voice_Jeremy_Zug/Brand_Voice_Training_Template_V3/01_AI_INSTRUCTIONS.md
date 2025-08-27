# AI Implementation Instructions for Brand Voice Guide Generation
Version 3.0 (Comprehensive, In-Depth & Actionable Focus)

## CORE OBJECTIVE: Generate a Comprehensive, In-Depth, and Actionable Brand Voice Guide

Your primary goal is to synthesize the provided brand materials into a **thorough, detailed, and highly illustrative** master brand voice guide. The aim is to create a robust resource comparable to best-in-class human-generated guides, rich with examples, explanations, and practical applications. While still actionable, the emphasis is on depth and completeness.

Concurrently, you will generate **detailed platform-specific add-on guides** that expand significantly on how the core brand voice adapts to unique platform contexts, including specific structural, tonal, and engagement nuances.

**Prioritize depth and clarity:** While conciseness in expressing individual points is good, do not sacrifice necessary detail, comprehensive explanations, or a rich array of examples. The final documents must provide a deep, actionable understanding of the brand's voice for both human strategists and consuming AIs.

---

## 1. Master Brand Voice Guide: Structure & Content

Create a single, consolidated `Generated_Brand_Voice_Master.md`. This document should synthesize information into the following core sections. **Do NOT create separate files for these sections initially; they are logical divisions within the single master document.**

   ```markdown
   # [Client Name] - Master Brand Voice Guide
   
   ## 00_How_To_Use_This_Guide_For_AI
   - Brief (1-2 paragraphs) instructions for the consuming AI.
   - Example: "You are [Brand Name]'s AI voice. Your primary goal is to generate content that is [key characteristic 1] while being [key characteristic 2]. Refer to the examples and guidelines below. Prioritize authenticity and direct application of these patterns."
   
   ## 01_Core_Brand_Identity
   - **Source Synthesis From:** processed client summaries in the `Summaries/` directory and relevant raw client content. **The AI MUST autonomously derive all core identity elements (Mission, Values, Personality, USP, Story) from these sources; it should NOT prompt the user to provide this information.**
   - **Content Focus (Emphasize depth and illustrative examples):**
     - Mission (1-2 clear, impactful sentences).
     - Core Values (3-5 bullet points, each with a brief explanatory sentence and, if possible, an example of how the value manifests in communication).
     - Brand Essence/Personality (Identify 3-5 primary descriptive adjectives. For each adjective: provide a detailed definition in the context of the brand, 2-3 distinct, verbatim examples from source material illustrating this trait, and a brief explanation of how this trait should be conveyed).
     - Unique Selling Proposition (1-2 concise, powerful statements).
     - Brand Story/Origin (If it *directly and significantly* informs the voice, provide a detailed narrative of 2-4 paragraphs, highlighting elements that shape communication style. Include illustrative quotes or anecdotes from source material if available).
   
   ## 02_Voice_Tone_And_Style
   - **Source Synthesis From:** processed client summaries and raw source materials, focusing on linguistic patterns, vocabulary, tone, stylistic choices, and recurring phrases.
   - **Content Focus (Expand significantly for depth and clarity):**
     - **Detailed Voice Characteristics:** Identify and detail 3-5 core voice characteristics (e.g., Warm, Authoritative, Witty, Empathetic, Playful). For each:
       - Provide a clear definition of what this characteristic means for the brand.
       - Offer 3-5 distinct, verbatim examples from source material that vividly illustrate this characteristic.
       - Briefly explain common applications or contexts where this characteristic is particularly important.
     - **Tone Spectrum & Mapping:** Describe the brand's primary tonal range (e.g., from Formal to Informal, Humorous to Serious, Technical to Accessible). 
       - Identify 3-5 key tonal variations or modes the brand uses (e.g., 'Educational & Empowering', 'Challenging & Provocative', 'Supportive & Understanding').
       - For each tonal variation, explain its purpose, typical contexts, and provide 2-3 examples of it in action. Include guidance on how to select the appropriate tone.
     - **Linguistic Patterns & Stylistic Choices:** Analyze and describe specific, recurring patterns:
       - Sentence Structure: Typical length (short, long, varied), complexity, common constructions (e.g., active vs. passive, use of questions).
       - Rhythm and Pacing: Is it fast-paced, measured, conversational, staccato? Provide examples.
       - Rhetorical Devices: Common use of metaphors, analogies, storytelling, humor, irony, etc., with examples.
       - Point of View: (e.g., first person, third person, collective 'we').
     - **Vocabulary & Lexicon (Brand Glossary):**
       - *Preferred Language:* A comprehensive list (15-20+ terms/phrases) of on-brand words, jargon (with explanations), and characteristic phrases. Provide context and 1-2 examples of use for each.
       - *Discouraged Language:* A list (10-15+ terms/phrases) of off-brand words, clichés to avoid, or terms that misrepresent the brand. Explain why they are discouraged and provide alternatives if applicable.
     - **Signature Phrases, Taglines & Hook Patterns:**
       - List all identified signature phrases and taglines with explanations of their meaning and use.
       - Document common or effective hook patterns used in content (e.g., starting with a question, a bold statement, a short anecdote). Provide 3-5 examples of different hook patterns.
     - **Do's and Don'ts:** Create a dedicated sub-section with at least 5-7 clear 'Do' statements and 5-7 'Don't' statements regarding voice and style. Each point must be accompanied by a specific, illustrative example (for 'Do's') or a counter-example (for 'Don'ts').
   
   ## 03_Content_Strategy_And_Audience
   - **Source Synthesis From:** processed client summaries and raw content related to content strategy, audience research, marketing plans, and engagement data.
   - **Content Focus (Expand for detail and voice application):**
     - **Core Content Themes/Pillars:** List 3-5 primary themes. For each theme:
       - Briefly describe the theme and its relevance to the brand and audience.
       - Provide 2-3 example content angles, headlines, or key messages that demonstrate how the brand voice applies to this theme.
     - **Detailed Audience Personas:** For each primary (and key secondary) audience segment:
       - Develop a brief persona including demographics (if available/relevant), psychographics, key needs, pain points, motivations, and communication preferences.
       - Describe how the brand voice should specifically adapt or resonate with this persona, providing examples or key considerations.
     - **General Approach to Content & Voice Application:** Articulate the overarching philosophy for content (e.g., "Educate and empower through practical, data-driven insights delivered with a confident and approachable voice."). Explain how the voice supports this approach.
     - **Key Calls to Action (General):** List 3-5 common general CTAs. For each, provide an example of how it's typically phrased in the brand's voice.
   
   ## 04_Key_Exemplars_And_Storytelling
   - **Source Synthesis From:** client-provided examples of 'gold standard' content, or AI-identified highly representative pieces from processed summaries/raw content.
   - **Content Focus (Emphasize detailed analysis and diverse examples):**
     - **Detailed Exemplar Analysis:** Select 3-5 diverse key exemplar texts (e.g., a blog post excerpt, an email segment, a social media post, a video script snippet). For each exemplar:
       - Present the verbatim text.
       - Provide a detailed analysis (2-3 paragraphs) breaking down *how* it embodies the brand voice. Refer explicitly to elements from `01_Core_Brand_Identity` and `02_Voice_Tone_And_Style` (e.g., "Notice the use of [preferred vocabulary] and the [specific tone characteristic] in the second sentence. The hook pattern aligns with [documented hook pattern X].").
       - Highlight specific word choices, sentence structures, and tonal qualities that make it effective.
     - **Storytelling Frameworks & Archetypes:**
       - If the brand utilizes specific storytelling frameworks (e.g., Hero's Journey, Problem-Agitate-Solve), describe them and provide examples of their application in the brand's content.
       - If core story archetypes (e.g., Sage, Explorer, Jester) are central to the brand, describe 1-3 key archetypes. For each, explain how it manifests in the voice and provide 2-3 illustrative examples.
   ```

---

## 2. Platform-Specific Add-On Guides: Structure & Content

After the Master Guide is drafted, you will be prompted to create **separate, detailed `.md` files** for each specified platform (e.g., `LinkedIn_Guide.md`, `Newsletter_Guide.md`). Aim for a comprehensive guide of **20-50 pages each**, to fully detail platform-specific voice application.

**Key Principles for Platform Guides:**
*   **Reference Master, Detail Nuances:** These guides should briefly reference core principles from the Master Guide but primarily focus on detailing the *specific adaptations, structures, examples, and engagement strategies* for that platform.
*   **Focus on In-Depth, Actionable Differences:** Go beyond brief notes. Provide rich, detailed explanations and multiple examples of what *specifically* changes or needs emphasis for this platform. This includes post structures, formatting, tone modulation, platform-specific vocabulary or jargon, visual integration with voice, and engagement best practices.

   **Template for Platform-Specific Guides:**
   ```markdown
   # [Platform Name] - Brand Voice Add-On
   
   **Reference:** This guide supplements the [Client Name] Master Brand Voice Guide. Refer to the Master Guide for core identity, voice, and style principles.
   
   ## 1. Purpose & Goals on [Platform Name]
   - What are the primary objectives for this brand on this platform? (1-2 sentences)
   
   ## 2. Content Structure & Format on [Platform Name]
   - Detail common post/content structures (e.g., LinkedIn: specific hook formulas, value proposition structure, multi-point arguments, CTA sequences; Newsletter: engaging subject line strategies, welcome section, main story structure, resource linking, closing remarks).
   - Provide 2-3 full examples of ideal post structures with annotations explaining each part.
   - Detail platform-specific formatting nuances (e.g., optimal use of emojis, hashtag strategy, video length and style, image and caption synergy, use of platform features like polls, Q&A, etc.), with examples.

   ## 3. Tone & Voice Nuances for [Platform Name]
   - Describe in detail any specific shifts or emphases in tone for [Platform Name] compared to the general brand voice (e.g., "On LinkedIn, adopt a more data-driven and professionally assertive tone, while maintaining core warmth. This means incorporating more statistics and industry insights, as shown in Example A and B."). Provide 2-3 illustrative examples of these tonal nuances.
   - List 3-5 platform-specific Do's and 3-5 Don'ts with clear examples for each, focusing on voice application unique to the platform.

   ## 4. Key Exemplar(s) for [Platform Name]
   - Provide 2-3 *detailed, specific examples* of successful content created by this brand *specifically for [Platform Name]*. These should clearly illustrate the voice, structure, and format adaptations for the platform.
   - For each exemplar, provide a brief analysis (1-2 paragraphs) similar to the Master Guide exemplar analysis, highlighting platform-specific voice application.

   ## 5. Calls to Action (CTAs) & Engagement for [Platform Name]
   - List 3-5 common and effective CTAs used by the brand on [Platform Name], with examples of how they are phrased in the brand voice.
   - Describe 2-3 key engagement strategies specific to this platform (e.g., "On Instagram, respond to comments within 4 hours using a friendly and appreciative tone, often asking a follow-up question as seen in Example C."), providing examples of voice in engagement.
   ```

---

## 3. AI Content Generation Rules (NEW)

*   **Overall Length (Master Guide):** Aim for a **comprehensive and in-depth guide of 20-50 pages**. Prioritize rich detail and illustrative examples within this range.
*   **Overall Length (Platform Guides):** Aim for **detailed and thorough platform-specific guides of 20-50 pages each**.
*   **Synthesis and Elaboration:** Do not simply copy-paste. Analyze, extract key concepts, and then elaborate with detailed explanations and rich examples to ensure deep understanding.
*   **Examples (CRITICAL):**
    *   Provide **multiple (3-5+) rich, illustrative examples** for each key point, principle, or characteristic described.
    *   Examples should be **verbatim** from source material whenever possible, or very closely paraphrased if necessary for clarity, always citing the source.
    *   Include **counter-examples ('What NOT To Do')** where they significantly clarify a voice principle or help avoid common pitfalls. Provide a brief explanation for why the counter-example is off-brand.
    *   Ensure examples are diverse and showcase a range of applications.
*   **Show, AND Tell Extensively:** Use abundant examples to *show* the voice in action. Accompany these examples with clear explanations (*tell*) that break down *why* the example is effective and *how* it applies the brand's voice principles.
*   **Minimize Unnecessary Redundancy, Maximize Clarity:** While core principles from the Master Guide shouldn't be repeated verbatim in platform guides, ensure each guide is self-contained enough to be understood. Cross-reference intelligently. The priority is clarity and depth within each section and guide.
*   **Traceability:** Add short inline citations for each quote, datapoint, or specific example, e.g. `(Source: Podcast_Ep174, 00:12:30)` or `(Source: Blog_Post_Title, para 3)` so future reviewers can locate the origin quickly.
*   **Clarity, Depth, Richness, and Actionability:** Every piece of information must be exceptionally clear, thoroughly explained, supported by rich examples, and directly actionable for both human users and consuming AIs.

---

## 4. Process & Quality

**Pre-Phase: Comprehensive Source Inventory (Critical for accurate progress tracking)**
0.  **Full Source File & Chunk Manifestation:** Before any ingestion begins (i.e., before summaries are created), the AI (or operator guiding the AI) MUST:
    a.  List ALL client source files provided (e.g., from `Raw/`, `Chunks/` or specified source folder).
    b.  For each file, determine the number of chunks required based on size and type (refer to `03_INGESTION_GUIDE.md` for chunking rules).
    c.  Populate the `Strategic_Context/Process_Tracker.md` with an entry for EVERY original file and EVERY anticipated chunk. This tracker becomes the master list defining the entire scope of the ingestion work.
    d.  The AI should not consider the ingestion phase complete, nor proceed to summary mapping, until all items in this pre-populated tracker are processed and marked as complete.

**Phase 0 – Mapping (High-Level Scan of *Summaries*)**
1.  Scan every file in the `Summaries/` directory to identify which raw sources contain information relevant to each guide section (mission, values, tone, exemplars, etc.). Make a quick index of *where the gold is*.
    *   If a summary lacks the needed fidelity (e.g., no direct quotes, vague references), flag that source for **re-processing** from the raw file.

**Phase 1 – Targeted Raw Dive**
2.  Open only the flagged raw files (or relevant portions) and pull **exact wording, illustrative quotes, stylistic patterns, and data points** required for the current section.
    *   Capture short, verbatim snippets; keep context as needed.

**Phase 2 – Drafting & Refinement**
3.  Draft the section in `Generated_Brand_Voice_Master.md`, using the extracted quotes with inline citations. Ensure brevity and high signal.
4.  Iterate: self-review for clarity, duplication, and length targets.

**Phase 3 – Platform Guides**
5.  After Master Guide approval, repeat the Mapping → Targeted Dive workflow for each platform-specific add-on.

**Quality Checklist**
6.  Each section must include: concise synthesis, at least one direct quote/example, citation, no redundancy, actionable guidance.
7.  Run a final pass to ensure tone consistency and correct citations.

---

## FINAL INSTRUCTION: Prioritize Depth, Richness, Clarity, and Actionability

Your main goal is to create training documents that are **exceptionally thorough, deeply illustrative, and highly effective** for both human strategists and consuming AIs. While clarity in language is vital, do not shy away from detail, extensive examples, or comprehensive explanations where they enhance understanding and application of the brand voice. The success of this system relies on your ability to synthesize complex brand information into a powerful, detailed, and actionable guide that leaves no room for ambiguity and fully captures the nuances of the brand's unique voice.
