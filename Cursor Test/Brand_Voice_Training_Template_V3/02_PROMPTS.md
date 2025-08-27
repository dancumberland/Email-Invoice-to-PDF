# Brand Voice Generation Prompts - V3
*Copy/Paste Prompts for AI-Assisted Brand Voice Guide Creation*

**📋 USAGE:** Copy each prompt below and paste into your AI conversation. Provide the client source materials as specified in each prompt.

**⚡ PREREQUISITE:** Ensure your AI has been initialized using the System Initialization Prompt from `00_START_HERE.md`

---
## I. Master Brand Voice Guide (`Generated_Brand_Voice_Master.md`)

### Prompt for: `00_How_To_Use_This_Guide_For_AI`

```
Generate a brief (1-2 paragraphs) set of instructions for an AI that will consume this brand voice guide. The instructions should clearly state:
1. The AI's role (e.g., "You are [Brand Name]'s AI voice.").
2. Its primary goal when generating content (e.g., "Your primary goal is to generate content that is [key characteristic 1] while being [key characteristic 2].").
3. A directive to refer to the examples and guidelines within this document.
4. An emphasis on prioritizing authenticity and direct application of the documented patterns.

Keep this section extremely concise and actionable for the consuming AI.
```

### Prompt for: `01_Core_Brand_Identity`

```
Synthesize information from the **collection of summary files in the `Summaries/` directory (generated from ingested raw assets like transcripts, posts, etc.)** to create the 'Core Brand Identity' section for the NEW Master Brand Voice Guide. The AI should look for recurring themes and explicit statements across all summaries to identify the brand's mission, values, personality, USP, and relevant story elements.

Extract and present the following, ensuring extreme conciseness and focus on high-signal information:

*   **Mission:** 1-2 concise sentences that capture the brand's core purpose.
*   **Core Values:** 3-5 key bullet points representing the brand's fundamental beliefs.
*   **Brand Essence/Personality:** 3-5 primary descriptive adjectives. For each adjective, provide a brief (1-sentence) explanation. If possible, include a very short, illustrative quote or phrase from the source material that exemplifies this trait.
*   **Unique Selling Proposition (USP):** 1 concise statement that clearly articulates what makes the brand different or special.
*   **Brief Brand Story/Origin (Conditional):** Include 1-2 short paragraphs (MAXIMUM) about the brand's story or origin ONLY IF these elements *directly and significantly* inform the brand's voice and communication style. If not directly voice-informing, omit this part or keep it to a single sentence.

Prioritize direct extraction of key phrases and concepts. Ruthlessly edit for brevity. The goal is a highly distilled summary of the brand's core identity as it pertains to voice.
```

### Prompt for: `02_Voice_Tone_And_Style`

```
Condense and synthesize information from the **"Voice Markers," "Notable Quotes," and "Summary" sections of all files within the `Summaries/` directory** to create the 'Voice, Tone, And Style' section for the NEW Master Brand Voice Guide. The AI should analyze these aggregated outputs from diverse raw materials to identify dominant voice characteristics, tonal spectrums, and preferred/discouraged language.

Focus on extracting and presenting the following with utmost brevity and impact:

*   **Overall Voice Characteristics:** Identify the 3-5 most dominant and distinctive voice characteristics (e.g., Warm, Authoritative, Witty, Empathetic, Playful). For each characteristic, provide 1-2 *brief, highly illustrative examples* directly quoted or closely paraphrased from the source material.
*   **Tone Spectrum:** List 2-3 key spectrums that define the brand's tonal range (e.g., Conversational yet Professional; Enthusiastic but Grounded). Provide a short explanation for each spectrum.
*   **Preferred Language/Vocabulary:** List 5-10 *key* preferred words or short phrases that are highly characteristic of the brand. For each, provide 1 brief example of its typical use from the source material. Focus on terms that are unique, frequently used, or particularly impactful for the brand's voice.
*   **Discouraged Language/Vocabulary:** List 5-10 *key* words or short phrases that the brand actively avoids. For each, provide 1 brief example or a concise reason for its avoidance (e.g., "Avoid: 'Synergy' - feels too corporate").
*   **Signature Phrases/Taglines:** If the brand has established signature phrases or taglines that are integral to its voice, list them.
*   **Sentence & Paragraph Style (Conditional):** Include brief notes on preferred sentence length, complexity, rhythm, or paragraph structure *only if these are distinctly characteristic and consistently applied*. If present, provide 1-2 illustrative examples from the source material. If not a strong differentiator, omit this part.

Be extremely selective. Prioritize the most impactful and representative elements that define the brand's unique sound and feel. Aim for a concise, actionable summary.
```

### Prompt for: `03_Content_Strategy_And_Audience`

```
Synthesize key insights from the **"Summary" and "Tags" sections of files in the `Summaries/` directory, and by analyzing the full content of ingested raw assets (from `Chunks/` or `Raw/` directories as appropriate)** to create the 'Content Strategy and Audience' section for the NEW Master Brand Voice Guide. The AI should infer primary audience segments, recurring content pillars/themes, and common calls to action by analyzing the actual communications of the brand.

Extract and present the following concisely:

*   **Core Content Themes/Pillars:** List 3-5 primary themes or content pillars that the brand consistently focuses on.
*   **Primary Audience Snapshot:** For each key audience segment identified in the source material, provide a brief description (2-3 sentences per segment). Focus on: Who they are, their key needs/motivations, and any specific voice considerations for communicating effectively with them.
*   **General Approach to Content:** A concise statement (1-2 sentences) summarizing the brand's overall philosophy or approach to content creation (e.g., "Educate and empower through practical storytelling," or "Provide actionable insights with a direct, no-nonsense approach that respects the audience's time.").
*   **Key Calls to Action (General):** List 2-3 common, general calls to action used by the brand (not platform-specific yet).

Ensure this section provides a high-level overview that informs voice application in content, without delving into platform-specific tactics.
```

### Prompt for: `04_Key_Exemplars_And_Storytelling`

```
Identify and extract 2-3 of the most potent and representative content exemplars by **reviewing the ingested raw content files (from `Chunks/` or `Raw/` directories) that have been highlighted as strong examples in their corresponding `Summaries/` files (e.g., via "Notable Quotes" or specific tags indicating exemplary quality).** For story archetypes, the AI should analyze recurring narrative patterns across multiple ingested raw assets and their summaries. This information will form the 'Key Exemplars And Storytelling' section for the NEW Master Brand Voice Guide.

Focus on:

*   **Showcase, Don't Just Describe:** Provide 2-4 *short, highly representative examples* (e.g., a well-chosen paragraph, a concise social media post snippet, a key excerpt from a 'About Us' page) of the brand's content that clearly demonstrate the voice characteristics outlined in section `02_Voice_Tone_And_Style`.
*   **Source of Exemplars:** Prioritize extracting these exemplars *directly* from the client's best existing materials. If absolutely necessary and no perfect short snippet exists, you may need to *craft a very short, illustrative example* that is deeply rooted in their documented patterns.
*   **Minimal Annotations:** For each exemplar, add a brief (1-sentence) annotation highlighting *why* it's a good example (e.g., "Note the use of [key voice characteristic] and [preferred phrase/metaphor from section 02].").
*   **Core Story Archetypes (Conditional):** If the brand has 1-2 truly central and consistently used story archetypes (e.g., "The Mentor's Journey," "The Underdog Triumph"), briefly describe each archetype (1-2 sentences) and provide one very short example of how it manifests in their content. If storytelling isn't a dominant, structured feature, omit this or keep it extremely brief.

The goal is to provide clear, tangible demonstrations of the voice in action, not theoretical descriptions.
```

---
## II. Platform-Specific Add-On Guides (e.g., `LinkedIn_Guide.md`, `Newsletter_Guide.md`)

### General Prompt Structure for ANY Platform-Specific Add-On Guide:

```
You are creating a [Platform Name] Brand Voice Add-On Guide for [Client Name]. This guide will be a **separate, lightweight .md file** (target: 3-5 pages MAX).

**CRITICAL:** This guide *supplements* the [Client Name] Master Brand Voice Guide. It should **reference** core principles from the Master Guide and only detail the *deltas* or specific applications for [Platform Name]. **DO NOT REPEAT** content already covered in the Master Guide.

Synthesize information by **primarily analyzing ingested raw content (from `Chunks/` or `Raw/`) and their corresponding `Summaries/` files that are specifically tagged or identified as relevant to [Platform Name] or serve as exemplars for it.** The AI should look for patterns in communication style, content structure, and engagement strategies demonstrated in these platform-specific materials.

Structure the [Platform Name]_Guide.md as follows:

# [Platform Name] - Brand Voice Add-On

**Reference:** This guide supplements the [Client Name] Master Brand Voice Guide. Refer to the Master Guide for core identity, voice, and style principles.

## 1. Purpose & Goals on [Platform Name]
   - Briefly state the primary objectives for this brand on [Platform Name] (1-2 sentences).

## 2. Content Structure & Format on [Platform Name]
   - Detail any specific post/content structures common for the brand on [Platform Name] (e.g., LinkedIn post: Hook, Value Points, CTA; Newsletter: Welcome, Main Story, Resource Links, Closing).
   - Note any platform-specific formatting nuances (e.g., use of emojis, hashtags, video length constraints, image styles).

## 3. Tone & Voice Nuances for [Platform Name]
   - Describe any specific shifts or emphases in tone for [Platform Name] compared to the general brand voice (e.g., "Slightly more informal and conversational on Instagram DMs," or "More professional and data-driven on LinkedIn articles.").
   - List 2-3 platform-specific Do's and Don'ts that are not covered in the Master Guide or need special emphasis here.

## 4. Key Exemplar(s) for [Platform Name]
   - Provide 1-2 *short, specific examples* of successful content created by this brand *specifically for [Platform Name]*. These should clearly illustrate the voice and format adaptations for the platform.
   - Add a brief (1-sentence) annotation to each example.

## 5. Calls to Action (CTAs) & Engagement for [Platform Name]
   - List 2-3 common and effective CTAs used by the brand on [Platform Name].
   - Briefly describe 1-2 key engagement strategies specific to this platform (e.g., "Respond to all comments within 24 hours," "Use polls to drive interaction in Stories.").

Ensure this guide is extremely concise, actionable, and focused *only* on the unique aspects of using the brand voice on [Platform Name].
```

**Instructions for AI using these prompts:**
*   Replace `[Client Name]` and `[Platform Name]` as appropriate.
*   When synthesizing information, the AI should primarily draw from the structured data in the **`Summaries/` directory**. For identifying broader context, full exemplars, or details not captured in summaries, the AI should refer to the **original ingested raw content files located in the `Chunks/` or `Raw/` directories**.
*   The process involves analyzing patterns, themes, and specific examples across multiple ingested source materials (transcripts, articles, posts, etc.) to build each section of the Brand Voice Guide.
*   Adhere strictly to the conciseness and example guidelines outlined in `01_AI_INSTRUCTIONS.md`.
