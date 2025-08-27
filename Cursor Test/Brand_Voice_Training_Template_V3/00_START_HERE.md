# Brand Voice Training Template V3 - START HERE

**🎯 OBJECTIVE:** Generate concise, AI-optimized brand voice guides (10-20 pages) plus robust platform add-ons (≈20 pages each) and strategic brand context documents from diverse client source materials.

---

## 📋 QUICK START CHECKLIST

### Phase 0: Source Ingestion (variable)
- [ ] **0.1** Review `03_INGESTION_GUIDE.md` for chunking rules
- [ ] **0.2** Place raw assets in `Raw/`, chunk oversized files into `Chunks/`
- [ ] **0.3** Run *Source Ingestion Prompt* (≤3 chunks at a time) → save summaries to `Summaries/`
- [ ] **0.4** Update `Strategic_Context/Process_Tracker.md` with each processed chunk

### Phase 1: Preparation (5-10 minutes)
- [ ] **1.1** Collect client source materials (transcripts, LinkedIn posts, newsletters, podcast interviews, training decks, etc.)
- [ ] **1.2** Create client folder: `/Brand_Voice_[ClientName]/`
- [ ] **1.3** Create subfolders: `Master_Guide/`, `Platform_Guides/`, and `Strategic_Context/`
- [ ] **1.4** Convert materials to text/markdown format if needed

### Phase 2: AI Setup (2 minutes)
- [ ] **2.1** Share the `01_AI_INSTRUCTIONS.md` file with your AI
- [ ] **2.2** Use the **System Initialization Prompt** below to start

### Phase 3: Master Guide Generation (15-30 minutes)
- [ ] **3.1** Generate each section using prompts from `02_PROMPTS.md`
- [ ] **3.2** Review and iterate on each section
- [ ] **3.3** Combine into final `Generated_Brand_Voice_Master.md`

### Phase 4: Platform Add-ons (30-60 minutes)
- [ ] **4.1** Generate robust platform-specific guides using `02_PROMPTS.md` (≈20 pages each)
- [ ] **4.2** Review for platform-specific accuracy and depth
- [ ] **4.3** Ensure comprehensive coverage of platform nuances

### Phase 5: Strategic Brand Context Documents (Optional - 45-90 minutes)
- [ ] **5.1** Generate North Star documents (vision, goals, purpose)
- [ ] **5.2** Create detailed ICP profiles 
- [ ] **5.3** Document complete offer suite
- [ ] **5.4** Develop strategy enablement guides (marketing, sales, email, etc.)

### Phase 6: Quality Control & Delivery
- [ ] **6.1** Final review using QC checklist below
- [ ] **6.2** Package and deliver to client

---

## 🚀 SYSTEM INITIALIZATION PROMPT

**Copy and paste this prompt to your AI to begin the brand voice generation process:**

```
BRAND VOICE & STRATEGIC CONTEXT GENERATION - SYSTEM INITIALIZATION

You are an expert brand strategist and voice analyst tasked with creating comprehensive brand documentation for a client.

CONTEXT:
- We're using the Brand Voice Training Template V3 system
- Core Goal: Create a 10-20 page master brand voice guide plus ≈20 page platform add-ons
- Extended Goal: Generate strategic brand context documents (North Star, ICPs, Offer Suite, Strategy Guides)
- Source materials will be diverse: transcripts, social posts, newsletters, interviews, training decks, etc.
- Output must be AI-optimized, business-actionable, and comprehensive

SYSTEM CAPABILITIES:
1. **Brand Voice Master Guide** - Core voice, tone, style, examples
2. **Platform-Specific Guides** - Deep-dive guides for LinkedIn, Newsletter, Website, etc. (≈20 pages each)
3. **Strategic Context Documents** - North Star vision, detailed ICPs, complete offer documentation, strategy enablement guides

YOU HAVE BEEN PROVIDED:
1. AI Instructions document (01_AI_INSTRUCTIONS.md) - Your technical guidelines
2. Prompt templates (02_PROMPTS.md) - Specific prompts for each section
3. Client source materials - [USER WILL SPECIFY]

PROCESS:
1. I'll provide you with client source materials
2. You'll use the section-specific prompts from 02_PROMPTS.md
3. We'll generate documents iteratively, starting with the Master Brand Voice Guide
4. Focus on synthesis, strategic depth, and high-signal examples
5. Each document should be immediately implementable

FIRST STEP: Please confirm you understand this expanded system and are ready to receive the client source materials for analysis.

Remember: Strategic depth with ruthless clarity. Every element must drive business results.
```

---

## 📊 QUALITY CONTROL CHECKLIST

Before delivery, verify each section meets these standards:

### Master Guide Requirements
- [ ] **Length:** 10-20 pages total (≈4-6k tokens)
- [ ] **Conciseness:** Every sentence adds value
- [ ] **Examples:** 1-3 concrete examples per key point
- [ ] **Voice-focused:** All content directly informs voice/tone decisions
- [ ] **AI-optimized:** Fits in single context window

### Platform Guide Requirements  
- [ ] **Length:** ≈20 pages each (comprehensive platform coverage)
- [ ] **Platform-specific:** Deep dive into unique platform constraints/opportunities
- [ ] **Additive:** References master guide, doesn't duplicate
- [ ] **Actionable:** Clear guidance for content creation
- [ ] **Comprehensive:** Covers strategy, tactics, best practices, examples

### Strategic Brand Context Requirements
- [ ] **Business-focused:** Aligns with broader business objectives
- [ ] **Implementable:** Clear action steps and frameworks
- [ ] **Scalable:** Grows with the business
- [ ] **Integrated:** Connects with brand voice and platform strategies

### Content Quality Standards
- [ ] **Clear hierarchy:** Logical section flow
- [ ] **Consistent formatting:** Standard markdown structure
- [ ] **Executable examples:** Real, usable content samples
- [ ] **Professional tone:** Client-ready presentation

---

## 🎯 DELIVERABLES STRUCTURE

### Master Guide Sections:
1. **How To Use This Guide For AI** - Setup instructions
2. **Core Brand Identity** - Mission, values, essence, USP
3. **Voice, Tone & Style** - Language patterns, preferences, examples
4. **Content Strategy & Audience** - Pillars, audience insights, engagement
5. **Key Exemplars & Storytelling** - Annotated best examples

### Platform Add-on Templates:
- **LinkedIn Guide** - Professional networking voice
- **Newsletter Guide** - Email communication style  
- **Website/Blog Guide** - Web content voice
- **[Custom Platform]** - As needed per client

### Strategic Brand Context Templates:
- **North Star Document** - Vision, goals, purpose
- **ICP Profile** - Detailed audience insights
- **Offer Suite Document** - Complete product/service offerings
- **Strategy Enablement Guide** - Marketing, sales, email, etc.

---

## 🎯 STRATEGIC BRAND CONTEXT DOCUMENTS (OPTIONAL EXPANSION)

Beyond voice and platform guides, this system can generate comprehensive business strategy documents that form the foundation for website creation, marketing strategy, business planning, and sales enablement.

### 📋 NORTH STAR DOCUMENTS
**Purpose:** Define long-term vision and goals that drive all business decisions
- **10-Year Vision:** Ultimate impact and legacy goals
- **5-Year Strategy:** Major milestones and market position
- **3-Year Objectives:** Concrete business targets and capabilities
- **1-Year Goals:** Specific, measurable outcomes
- **Core Purpose ("The Why"):** Fundamental reason for existence

### 👥 ICP (IDEAL CUSTOMER PROFILE) SUITE
**Purpose:** Deep audience understanding for targeted marketing and product development
- **Primary ICP:** Main revenue-driving customer segment
- **Secondary ICPs:** Additional valuable segments
- **Anti-ICP:** Who to avoid targeting
- **Customer Journey Mapping:** Touchpoints and decision factors
- **Pain Points & Motivations:** Psychological and practical drivers

### 💼 COMPREHENSIVE OFFER SUITE DOCUMENTATION
**Purpose:** Complete inventory of products/services for sales and marketing alignment
- **Core Offerings:** Primary products/services
- **Pricing Strategy:** Value-based pricing framework
- **Product Ladder:** Customer journey through offerings
- **Bundling Options:** Package combinations
- **Upsells/Cross-sells:** Revenue expansion opportunities

### 🚀 STRATEGY ENABLEMENT GUIDES
**Purpose:** Actionable frameworks for business execution
- **Marketing Strategy Guide:** Channel strategy, content themes, campaign frameworks
- **Sales Enablement Guide:** Objection handling, sales scripts, process optimization
- **Email Strategy Guide:** Nurture sequences, segmentation, automation
- **Website Strategy Guide:** User experience, conversion optimization, content architecture
- **Business Development Guide:** Partnership strategy, growth initiatives

---

## 🔧 FILE REFERENCE

- **`01_AI_INSTRUCTIONS.md`** - Technical instructions for AI systems
- **`02_PROMPTS.md`** - Copy/paste prompts for each section
- **`03_INGESTION_GUIDE.md`** - Chunking workflow & ingestion prompt
- **`04_PROCESS_TRACKER_TEMPLATE.md`** - Copy to client folder for live tracking

---

## 📝 PROCESS NOTES

**Source Material Types:** Transcripts, training materials, presentation decks, podcast interviews, newsletters, LinkedIn posts, website copy, etc.

**Key Principle:** Synthesis over creation. Extract and distill existing voice patterns rather than inventing new ones.

**Success Metric:** A new AI can load the master guide and immediately produce on-brand content that sounds authentically like the client.

---

*Last Updated: 2025-06-03 | Version: 3.0*
