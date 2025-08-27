# 🗄️ Database Setup Guide

Your Claude API key is configured! Now let's set up the databases (both have generous free tiers).

## 📊 Option 1: Vercel Postgres (Recommended)

**Why Vercel Postgres:**
- ✅ 60 hours of compute per month (free)
- ✅ 256 MB storage
- ✅ Easy setup and management
- ✅ Perfect for development and light production use

**Setup Steps:**
1. Go to: https://vercel.com/dashboard
2. Click "Storage" → "Create Database" → "Postgres"
3. Name it: `inbox-zero-db`
4. Copy the connection strings to your clipboard
5. Update your `.env` file with the provided values:

```bash
# Edit the environment file
nano inbox-zero/apps/web/.env

# Replace these lines with your Vercel Postgres values:
POSTGRES_URL="your-vercel-postgres-url"
POSTGRES_PRISMA_URL="your-vercel-prisma-url"
POSTGRES_URL_NO_SSL="your-vercel-no-ssl-url"
POSTGRES_URL_NON_POOLING="your-vercel-non-pooling-url"
```

## 🔴 Option 2: Upstash Redis

**Why Upstash Redis:**
- ✅ 10,000 requests per day (free)
- ✅ Global edge locations
- ✅ Perfect for caching and session management

**Setup Steps:**
1. Go to: https://upstash.com/
2. Sign up/login
3. Click "Create Database"
4. Name it: `inbox-zero-redis`
5. Choose a region close to you
6. Copy the connection details
7. Update your `.env` file:

```bash
# Add to your .env file:
UPSTASH_REDIS_URL="your-upstash-redis-url"
UPSTASH_REDIS_TOKEN="your-upstash-redis-token"
```

## 🚀 Quick Launch Script

Once both databases are configured, run:

```bash
./start_inbox_zero.sh
```

Then open: http://localhost:3000

## 🎯 What Happens Next

1. **First Login**: You'll authenticate with your Gmail account
2. **AI Configuration**: Claude will be ready to process your emails
3. **Label Integration**: Your perfect label structure will be available
4. **Email Triage**: Claude will start learning your email patterns

## 🤖 Initial AI Rules to Set Up

Once logged in, create these rules:

```
Rule 1: HARO Emails
If from "haro@helpareporter.com":
→ Add label: @Reference/Inboxes/HARO
→ Archive email

Rule 2: Indy Acquisitions
If from domain "indyacquisitions.com":
→ Add label: @Reference/Inboxes/Indy Acquisitions
→ Archive email

Rule 3: Scheduling Requests
If subject contains ["schedule", "calendar", "meeting", "call"]:
→ Add label: @Action ⚡️/Schedule
→ Keep in inbox

Rule 4: Prospect Follow-up
If from prospect and no reply in 10 days:
→ Add label: @Prospect 💰/Hot-10d
→ Keep in inbox for follow-up

Rule 5: Action Items
If email needs response and is important:
→ Add label: @Action ⚡️/Respond
→ Keep in inbox
```

## 🔧 Troubleshooting

**Database Connection Issues:**
- Double-check connection strings in .env
- Ensure no extra spaces or quotes
- Restart the application after changes

**Claude API Issues:**
- Verify API key is correct
- Check Anthropic console for usage limits
- Ensure sufficient credits in your account

Ready to launch your AI email assistant! 🎉
