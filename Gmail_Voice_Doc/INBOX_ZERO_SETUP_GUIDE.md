# 🤖 Inbox Zero AI Rules Setup Guide

Your advanced AI email automation rules are ready to implement! Since you already have Gmail filters handling HARO and Indy Acquisitions, these rules focus on intelligent tasks that only AI can handle.

## 🎯 **Quick Setup Instructions**

### **Step 1: Access Your Inbox Zero Dashboard**
1. Go to: http://localhost:3000
2. Sign in with your Google account
3. Navigate to **"Rules"** or **"AI Assistant"** section

### **Step 2: Set Up These 6 Advanced AI Rules**

---

## **Rule 1: Hot Prospect Follow-up** 🔥
**Purpose:** Automatically manage prospect follow-ups and draft personalized responses

**Setup in Inbox Zero:**
```
Rule Name: Hot Prospect Follow-up
Trigger: When email is from a prospect AND no reply from me in 7+ days
Actions:
- Add label: @Prospect 💰/Hot-10d
- Draft personalized follow-up response
- Set reminder for 3 days
- Extract: company, project scope, budget mentions, timeline

Instructions for Claude:
"When I haven't replied to a prospect in 7+ days, add the @Prospect 💰/Hot-10d label and draft a personalized follow-up. Use Dan's authentic voice from Dan_Email_Voice_Training_MASTER.md: For existing contacts, start with 'Hey [Name]! How are things?!?' For new prospects, use 'Hi [Name],' with measured enthusiasm. Match enthusiasm level to relationship depth. Reference specific details from our previous conversation. Include 'Would love to connect' and suggest next steps. End with '-Dan'"
```

---

## **Rule 2: Urgent Action Detection** ⚡
**Purpose:** Identify and prioritize emails requiring immediate response

**Setup in Inbox Zero:**
```
Rule Name: Urgent Action Detection
Trigger: Email contains urgency indicators (urgent, asap, deadline, time-sensitive, emergency) AND requires response
Actions:
- Add label: @Action ⚡️/Respond
- Analyze priority level
- Draft urgent acknowledgment response
- Flag if from client domain

Instructions for Claude:
"When an email contains urgency indicators and needs a response, add @Action ⚡️/Respond label. Draft acknowledgment using Dan's voice from Dan_Email_Voice_Training_MASTER.md: Start with 'Hi [Name],' show understanding of urgency, provide specific timeline, include 'I apologize for any trouble this may cause' if appropriate. Keep paragraphs short (2-4 sentences). End with '-Dan'"
```

---

## **Rule 3: Smart Scheduling Assistant** 📅
**Purpose:** Handle scheduling requests with intelligent calendar awareness

**Setup in Inbox Zero:**
```
Rule Name: Smart Scheduling Assistant
Trigger: Email mentions scheduling, meeting, call, calendar, available, book time
Actions:
- Add label: @Action ⚡️/Schedule
- Extract scheduling details (times, duration, attendees)
- Draft scheduling response with available options
- Include calendar booking link if appropriate

Instructions for Claude:
"For scheduling requests, add @Action ⚡️/Schedule label and draft response using Dan's authentic voice from Dan_Email_Voice_Training_MASTER.md: For existing contacts, start with 'Hey [Name]!' For new contacts, use 'Hi [Name],' Match enthusiasm to relationship depth. Use 'Absolutely!' for enthusiasm. Reference specific meeting purpose. Use phrases like 'Would love to connect about [topic]' and 'Looking forward to it!' Include appropriate booking link from context file. End with '-Dan'"
```

---

## **Rule 4: Project Context Manager** 📋
**Purpose:** Intelligently route and contextualize project-related emails

**Setup in Inbox Zero:**
```
Rule Name: Project Email Router
Trigger: Email relates to active projects (MRR Method, Networking, REIT, AI Cohort)
Actions:
- Analyze project relevance
- Add appropriate @Projects/[ProjectName] label
- Extract action items and deadlines
- Update project status if mentioned

Instructions for Claude:
"When an email relates to one of my active projects (MRR Method, Networking, REIT, AI Cohort), analyze the content and add the appropriate @Projects/[ProjectName] label. Extract any action items, deadlines, or status updates. If the email mentions project completion, delays, or milestones, note this for project tracking."
```

---

## **Rule 5: Client Communication Enhancement** 👥
**Purpose:** Enhance client communications with context and voice consistency

**Setup in Inbox Zero:**
```
Rule Name: Client Email Intelligence
Trigger: Email is from a client domain or known client contact
Actions:
- Maintain conversation context from previous emails
- Draft response in appropriate client communication style
- Flag if requires immediate attention (complaints, urgent requests, payment issues)
- Track satisfaction indicators

Instructions for Claude:
"For client emails, maintain context from our conversation history. Draft responses in my professional, warm client communication style. Reference specific project details when relevant. Flag immediately if the email indicates dissatisfaction, urgent requests, or payment issues. Always end with clear next steps."
```

---

## **Rule 6: Revenue Opportunity Scanner** 💰
**Purpose:** Identify and prioritize potential business opportunities

**Setup in Inbox Zero:**
```
Rule Name: Revenue Opportunity Scanner
Trigger: Email mentions budget, timeline, project scope, hire, consulting, proposal (from non-existing clients)
Actions:
- Add label: @Prospect 💰/Hot-10d
- Extract opportunity details (budget, timeline, project type)
- Generate qualification questions
- Set follow-up sequence
- Calculate opportunity score

Instructions for Claude:
"When someone mentions budget, timeline, project scope, or hiring needs, this is a potential opportunity. Add @Prospect 💰/Hot-10d label and extract key details: budget range, timeline, project type, decision maker. Draft qualification questions to better understand their needs. Use my consultative, value-focused communication style."
```

---

## **🎯 Voice Training Instructions for All Rules**

Add this to each rule's instructions:

```
Voice Guidelines (Reference Dan_Email_Voice_Training_MASTER.md):
- Authentic warmth with professional competence
- **Match enthusiasm to relationship depth:** Full enthusiasm ("How are things?!?") for existing contacts, measured enthusiasm ("Hi [Name],") for new prospects
- Use conversational tone with contractions (89% of emails)
- Keep paragraphs short (2-4 sentences max)
- Share relevant personal context when appropriate
- Offer help generously: "Anything I can help with?"
- End with simple "-Dan" signature
- Prioritize relationship over transaction
```

---

## **⚙️ Advanced Settings**

### **Processing Schedule:**
- Set rules to run every 15 minutes
- Batch process up to 50 emails at a time
- Confidence threshold: 75%

### **Performance Monitoring:**
- Track rule effectiveness
- Measure time saved
- Monitor response quality
- Generate weekly summary reports

---

## **🚀 Implementation Priority**

**Start with these 3 rules first:**
1. **Hot Prospect Follow-up** (highest ROI)
2. **Urgent Action Detection** (prevents missed opportunities)
3. **Smart Scheduling Assistant** (saves most time)

**Then add:**
4. Project Context Manager
5. Client Communication Enhancement
6. Revenue Opportunity Scanner

---

## **📊 Expected Results**

With these rules active, you should see:
- **50-70% reduction** in manual email triage time
- **Faster response times** to prospects and clients
- **Better organization** with automatic labeling
- **Consistent authentic voice** matching your Dan_Email_Voice_Training_MASTER.md patterns
- **No missed opportunities** due to automated follow-ups
- **Relationship-first communication** that builds genuine connections

---

## **🔧 Troubleshooting**

**If a rule isn't working:**
1. Check the trigger conditions are specific enough
2. Verify Claude has access to your Gmail labels
3. Test with a sample email first
4. Adjust confidence threshold if needed

**For client demos:**
- Show before/after email processing
- Demonstrate voice consistency
- Highlight time savings metrics
- Show intelligent prioritization in action

---

## **📈 Next Steps**

1. **Implement the first 3 rules** in Inbox Zero
2. **Test with recent emails** to verify functionality
3. **Monitor performance** for first week
4. **Adjust rules** based on results
5. **Add remaining rules** once first 3 are optimized
6. **Document results** for client presentations

Your intelligent email automation system is now ready to handle the sophisticated tasks that Gmail filters can't manage! 🎉
