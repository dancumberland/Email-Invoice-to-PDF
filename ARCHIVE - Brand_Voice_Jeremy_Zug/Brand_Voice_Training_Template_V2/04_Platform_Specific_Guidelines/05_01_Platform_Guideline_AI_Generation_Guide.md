# 05_01 Platform Guideline AI Generation Guide

## Objective for AI

Your primary objective is to analyze all provided client materials and generate distinct, detailed, and actionable voice, formatting, syntax, and strategic guidelines for each relevant marketing platform the client uses or intends to use. Each generated guideline document should be a self-contained markdown file.

## Core Principles for AI

1.  **Client-Data Centricity:** All generated guidelines MUST be rooted in, and directly reference, client-provided materials (e.g., existing content, brand identity documents, ICPs, marketing strategies, style guides).
2.  **Specificity and Actionability:** Guidelines should be precise and provide clear, actionable instructions.
3.  **Platform Nuance:** Recognize and detail the unique conventions and best practices for each platform, as reflected in the client's data or, if necessary, inferred by adapting general guidelines to platform norms.
4.  **Transparency in Sourcing:** Clearly indicate the source of each piece of information. If specific data for a point is unavailable for a given platform, note this. If inferences are made (e.g., adapting general guidelines because no platform-specific examples exist), clearly state the basis for the inference.
5.  **Dynamic Generation:** The number and type of platform guideline documents produced will depend entirely on the platforms identified as relevant from the client's materials.

## Inputs for AI Analysis

To perform this task, you must meticulously analyze:

*   All documents within the `01_Core_Brand_Identity/` directory.
*   All documents within the `02_ICPs/` directory (especially for identifying target audience platforms and communication preferences).
*   All documents within the `03_General_Content_Guidelines/` directory.
*   Any client-provided examples of their content on various platforms (e.g., spreadsheets, links, document dumps).
*   Any explicit list from the client detailing platforms they currently use or plan to target.

## AI Process for Guideline Generation

### Step 1: Platform Identification & Prioritization

1.  **Analyze Inputs:** Review all provided client materials.
2.  **Identify Platforms:** Compile a list of all distinct marketing/communication platforms the client currently uses, has used, or expresses a clear intention to use. Note where ICP analysis suggests a platform might be relevant even if not explicitly mentioned by the client.
3.  **Source Mapping:** For each identified platform, briefly note the primary client materials that provide information about its use (e.g., "Client's LinkedIn company page URL provided," "Newsletter examples attached," "ICP interviews mention Instagram as a key channel").
4.  **Prioritization (If Necessary):** If a very large number of potential platforms are identified, note if the client has indicated any priority platforms.

### Step 2: Guideline Document Generation (Repeat for Each Identified Platform)

For each platform identified in Step 1, create a new markdown file named `05_XX_[PlatformName]_Guidelines.md` (e.g., `05_02_LinkedIn_Guidelines.md`, `05_03_Instagram_Guidelines.md`, etc., incrementing `XX` sequentially starting from `02`).

Use the following structure and prompts to populate each file. **You, the AI, are to fill in the details based on your analysis of the client's materials.**

```markdown
# [Platform Name] - Voice, Formatting & Strategic Guidelines

**Platform:** [Automatically insert the specific platform name here, e.g., LinkedIn, Instagram, Client Newsletter, Company Blog]

**Date Generated:** [Insert current date]

**Source Materials Summary:** [AI, briefly list the key client documents/data points that informed this specific platform guide, e.g., "Analysis of 50 recent Instagram posts provided by client (Client_Instagram_Data.zip), 01_01_Brand_Story.md, 02_01_ICP_Development_AI_Guide.md for Segment A preferences."]

## 1. Overview & Strategic Purpose on [Platform Name]

*   **Brand's Primary Goals on [Platform Name]:**
    *   AI: Based on client materials (explicit statements, content themes, CTAs used), what are the 1-3 primary strategic goals for the brand on this platform? (e.g., Lead generation, Brand awareness, Community building, Direct sales, Thought leadership).
*   **Target Audience Segment(s) on [Platform Name]:**
    *   AI: Which ICP(s) are primarily targeted on this platform? What does client data reveal about their expectations or behaviors here?

## 2. Voice & Tone on [Platform Name]

*   **Core Voice Characteristics (Platform-Adjusted):**
    *   AI: Referencing `03_01_Brand_Voice_And_Tone.md`, how are the core voice attributes (e.g., 'Authoritative but Accessible,' 'Playful and Witty') specifically manifested or nuanced on [Platform Name]? Provide examples from client content if available, or explain the adaptation.
*   **Language & Terminology:**
    *   AI: Are there specific terms, jargon, or phrases the client uses (or avoids) on this platform? Refer to `03_04_Vocabulary_And_Terminology_Guide.md` and platform-specific examples.
*   **Emotional Resonance:**
    *   AI: What is the desired emotional impact of content on this platform, according to client data (e.g., inspiring, reassuring, exciting)?

## 3. Formatting & Syntax Rules for [Platform Name]

*   **Line Length & Paragraphs:**
    *   AI: Analyze client's [Platform Name] content. Describe typical/max line lengths and paragraph structure (e.g., "Prefers short, 1-2 sentence paragraphs for scannability on mobile," or "Uses longer, more detailed paragraphs in blog posts"). Note use of line breaks.
*   **Spacing:**
    *   AI: Detail conventions for single vs. double line breaks between paragraphs, around headings, etc., as observed in client content or platform best practices if client data is sparse.
*   **Headings/Subheadings (if applicable):**
    *   AI: How are sections structured? (e.g., "Uses bolding for informal section breaks in LinkedIn posts," "Follows H2, H3 markdown for blog structure").
*   **Lists (Bullet Points/Numbered):**
    *   AI: Describe the style and usage of lists (e.g., "Uses asterisk bullets for benefits," "Numbered lists for step-by-step instructions").
*   **Emphasis (Bold, Italics):**
    *   AI: How and when are bolding or italics used for emphasis? (e.g., "Bolds key takeaways," "Italicizes calls to action"). Provide examples if possible.
*   **Links & CTAs:**
    *   AI: How are links formatted and presented? What are common Call to Action phrasings and their placement/formatting? (e.g., "Directly pastes URLs in newsletters," "Uses 'Learn More -> [link]' in social posts"). Note any UTM tagging conventions mentioned.
*   **Hashtags (if applicable):**
    *   AI: What is the client's hashtag strategy on this platform? (e.g., Number to use, placement - in-line or at end, specific brand/campaign hashtags, mix of broad and niche). Refer to `03_06_Hashtag_Strategy_Guide.md` if it exists, and analyze actual usage.
*   **Emojis (if applicable):**
    *   AI: What is the client's approach to emojis on this platform? (e.g., "Permitted and frequently used to add personality," "Used sparingly and only specific ones," "Not used"). List preferred or commonly used emojis if identifiable.

## 4. Content Length & Cadence for [Platform Name]

*   **Typical Post/Content Length:**
    *   AI: Analyze client examples. What are the typical lengths (e.g., word count, character count, video duration) for primary content types on this platform?
*   **Posting Frequency & Timing (if data available):**
    *   AI: Does client data indicate a typical posting frequency (e.g., 3 times/week) or preferred times/days for this platform? If not, note as an area for client input.

## 5. Visuals on [Platform Name] (if applicable)

*   **Image/Video Style:**
    *   AI: Based on client examples or `01_04_Visual_Identity_Elements.md`, describe the aesthetic and style guidelines for images, graphics, or video thumbnails used on this platform.
*   **Alt Text Requirements:**
    *   AI: What are the client's practices or stated requirements for image alt text to ensure accessibility?

## 6. Platform-Specific Nuances & Features for [Platform Name]

*   **Feature Utilization:**
    *   AI: Does the client utilize platform-specific features (e.g., Instagram Stories/Reels, LinkedIn Polls/Articles, Twitter Threads, Facebook Groups, YouTube Shorts, Podcast show notes conventions)? Describe how these are used, or if not used, whether ICP/Strategy suggests they should be considered.
*   **Engagement & Interaction Style:**
    *   AI: How does the client typically engage with their audience on this platform? (e.g., "Responds to all comments within 24 hours with a friendly tone," "Uses questions in posts to encourage interaction").

## 7. Key Performance Indicators (KPIs) for [Platform Name] (if data available)

*   **Success Metrics:**
    *   AI: Does client data mention specific KPIs they track for this platform? (e.g., Engagement rate, Click-through rate, Follower growth, Conversions).

## 8. Points for Client Clarification / Data Gaps

*   AI: List any specific aspects of [Platform Name] strategy, formatting, or voice where client data was sparse or ambiguous, and where direct client input would be beneficial for refining these guidelines.

```

### Step 3: Final Review & Output

1.  **Consistency Check:** Ensure all generated documents are consistent with the core brand identity and general content guidelines, adapting them appropriately for each platform.
2.  **Completeness:** Verify that all sections in the template above have been addressed for each platform, noting where data was unavailable.
3.  **Output:** Provide the collection of generated markdown files.

## Example of AI Self-Correction/Refinement during the process:

*   "Initial analysis of LinkedIn posts showed inconsistent use of emojis. However, the `03_01_Brand_Voice_And_Tone.md` guide emphasizes a 'professional yet modern' voice. Cross-referencing with ICP data for LinkedIn (professionals aged 30-50), I will recommend minimal, contextually relevant emoji use, such as ✅ or ➡️, rather than more playful ones, unless client examples strongly contradict this. I will note this as an inference in the LinkedIn guide."
*   "No specific data on newsletter paragraph length was found. Based on general readability best practices for email and the client's stated goal of 'clear communication' in `01_03_Brand_Goals_And_Vision.md`, I will recommend short to medium paragraphs (3-5 sentences) and will flag this for client confirmation."
