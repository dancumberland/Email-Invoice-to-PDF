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
