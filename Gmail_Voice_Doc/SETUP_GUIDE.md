# 🎉 Inbox Zero Setup Complete!

## ✅ What's Ready:
- Repository cloned and dependencies installed
- Environment file created with secure secrets  
- Google OAuth credentials configured
- Startup script created

## 🔑 REQUIRED: Add OpenAI API Key

**Before starting, you MUST:**
1. Get API key from: https://platform.openai.com/api-keys
2. Edit `inbox-zero/apps/web/.env`
3. Replace: `OPENAI_API_KEY=""`
4. With: `OPENAI_API_KEY="sk-your-actual-key"`

## 🗄️ REQUIRED: Set Up Database & Redis

You need these cloud services (both have free tiers):

### PostgreSQL (Vercel Postgres - Recommended)
1. Go to: https://vercel.com/dashboard
2. Create new project → Storage → Postgres
3. Copy all POSTGRES_* variables to your .env file

### Redis (Upstash - Recommended)  
1. Go to: https://upstash.com/
2. Create Redis database
3. Copy UPSTASH_REDIS_URL and UPSTASH_REDIS_TOKEN to .env

## 🚀 Start Inbox Zero

```bash
./start_inbox_zero.sh
```

Then open: http://localhost:3000

## 🎯 Your Perfect Gmail Labels

Your labels are already organized for AI automation:

**Action Buckets:**
- @Action ⚡️/Respond
- @Action ⚡️/Todo  
- @Action ⚡️/Schedule

**Prospect Management:**
- @Prospect 💰/Hot-10d

**Waiting/Reminders:**
- @Waiting ⏳/7d, @Waiting ⏳/14d

**Reference:**
- @Reference/Inboxes/ (AI, HARO, News, Indy Acquisitions)
- @Reference/Courses/, @Reference/Swipes/

**Projects & Archive:**
- @Projects/, @Archive/Projects/, @Archive/Misc/

## 🤖 AI Rules Examples

Once running, you can set up rules like:

```
If from "haro@helpareporter.com":
→ Label: @Reference/Inboxes/HARO
→ Archive

If subject contains "schedule":
→ Label: @Action ⚡️/Schedule
→ Keep in inbox

If from prospect and no reply in 10 days:
→ Label: @Prospect 💰/Hot-10d
→ Keep for follow-up
```

## 🔧 Troubleshooting

**"OpenAI API key not set"**: Add your key to .env file
**Database errors**: Set up Vercel Postgres and update .env
**Redis errors**: Set up Upstash Redis and update .env

Happy email automation! 🎉
