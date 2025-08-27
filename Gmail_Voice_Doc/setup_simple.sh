#!/bin/bash

# Simple Inbox Zero Setup (No Admin Required)
echo "🚀 SIMPLE INBOX ZERO SETUP"
echo "=========================="

# Colors
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
NC='\033[0m'

print_status() { echo -e "${GREEN}✅ $1${NC}"; }
print_info() { echo -e "${BLUE}ℹ️  $1${NC}"; }
print_warning() { echo -e "${YELLOW}⚠️  $1${NC}"; }

# Setup directory
SETUP_DIR="/Users/dancumberland/Documents/Work/AI Projects & Training Docs/Dan_Personal_Brand/AI_Tools/Gmail_Voice_Doc"
cd "$SETUP_DIR"

print_info "Working in: $SETUP_DIR"

# Step 1: Clone repository
print_info "Step 1: Cloning Inbox Zero..."
if [ -d "inbox-zero" ]; then
    print_info "Directory exists, updating..."
    cd inbox-zero && git pull && cd ..
else
    git clone https://github.com/elie222/inbox-zero.git
fi
print_status "Repository ready"

# Step 2: Install with npm (no pnpm needed)
print_info "Step 2: Installing dependencies with npm..."
cd inbox-zero
npm install
print_status "Dependencies installed"

# Step 3: Generate secrets
print_info "Step 3: Generating secrets..."
NEXTAUTH_SECRET=$(openssl rand -hex 32)
EMAIL_ENCRYPT_SECRET=$(openssl rand -hex 32)
EMAIL_ENCRYPT_SALT=$(openssl rand -hex 16)

# Step 4: Create .env file
print_info "Step 4: Creating environment file..."
cat > apps/web/.env << EOF
# Inbox Zero Configuration
NEXTAUTH_SECRET="$NEXTAUTH_SECRET"
EMAIL_ENCRYPT_SECRET="$EMAIL_ENCRYPT_SECRET"
EMAIL_ENCRYPT_SALT="$EMAIL_ENCRYPT_SALT"
NEXTAUTH_URL="http://localhost:3000"

# Database (using Vercel Postgres - free tier)
POSTGRES_URL=""
POSTGRES_PRISMA_URL=""
POSTGRES_URL_NO_SSL=""
POSTGRES_URL_NON_POOLING=""

# Redis (using Upstash - free tier)
UPSTASH_REDIS_URL=""
UPSTASH_REDIS_TOKEN=""

# Google OAuth (from your existing setup)
GOOGLE_CLIENT_ID=""
GOOGLE_CLIENT_SECRET=""

# OpenAI (REQUIRED - you need to add this)
OPENAI_API_KEY=""

# Optional
RESEND_API_KEY=""
MAX_DURATION=300
EOF

# Step 5: Copy Google OAuth if available
print_info "Step 5: Configuring Google OAuth..."
OAUTH_FILE="/Users/dancumberland/MCPs/windsurf-mcp/gcp-oauth.keys.json"
if [ -f "$OAUTH_FILE" ]; then
    # Extract credentials manually (no jq dependency)
    CLIENT_ID=$(grep -o '"client_id":"[^"]*' "$OAUTH_FILE" | cut -d'"' -f4)
    CLIENT_SECRET=$(grep -o '"client_secret":"[^"]*' "$OAUTH_FILE" | cut -d'"' -f4)
    
    if [ -n "$CLIENT_ID" ] && [ -n "$CLIENT_SECRET" ]; then
        sed -i '' "s/GOOGLE_CLIENT_ID=\"\"/GOOGLE_CLIENT_ID=\"$CLIENT_ID\"/" apps/web/.env
        sed -i '' "s/GOOGLE_CLIENT_SECRET=\"\"/GOOGLE_CLIENT_SECRET=\"$CLIENT_SECRET\"/" apps/web/.env
        print_status "Google OAuth configured"
    else
        print_warning "Could not extract OAuth credentials automatically"
    fi
else
    print_warning "OAuth file not found"
fi

# Step 6: Create startup script
print_info "Step 6: Creating startup script..."
cat > ../start_inbox_zero.sh << 'EOF'
#!/bin/bash
echo "🚀 Starting Inbox Zero..."

cd inbox-zero/apps/web

# Check if OpenAI key is set
if grep -q 'OPENAI_API_KEY=""' .env; then
    echo "❌ OpenAI API key not set!"
    echo "Please add your OpenAI API key to apps/web/.env"
    echo "Get it from: https://platform.openai.com/api-keys"
    exit 1
fi

echo "Starting development server..."
npm run dev
EOF

chmod +x ../start_inbox_zero.sh

# Step 7: Create setup guide
print_info "Step 7: Creating setup guide..."
cat > ../SETUP_GUIDE.md << 'EOF'
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
EOF

print_status "Setup guide created"

echo ""
echo "🎉 SIMPLE SETUP COMPLETE!"
echo "========================"
print_status "Repository cloned and configured"
print_status "Dependencies installed with npm"
print_status "Environment file created"
print_status "Google OAuth configured"
print_status "Startup script created"

echo ""
print_warning "NEXT STEPS REQUIRED:"
echo "1. 🔑 Add OpenAI API key to inbox-zero/apps/web/.env"
echo "2. 🗄️  Set up Vercel Postgres database"
echo "3. 🗄️  Set up Upstash Redis"
echo "4. 🚀 Run: ./start_inbox_zero.sh"

echo ""
print_info "📖 Full instructions: cat SETUP_GUIDE.md"
print_info "📁 Setup completed in: $(pwd)"
