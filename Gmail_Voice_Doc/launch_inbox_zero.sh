#!/bin/bash

# Inbox Zero Automated Launch Script
# Handles Node.js version management and dependency installation

set -e

echo "🚀 LAUNCHING INBOX ZERO AI EMAIL ASSISTANT"
echo "=========================================="

# Colors
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
NC='\033[0m'

print_status() { echo -e "${GREEN}✅ $1${NC}"; }
print_info() { echo -e "${BLUE}ℹ️  $1${NC}"; }
print_warning() { echo -e "${YELLOW}⚠️  $1${NC}"; }

# Load NVM
export NVM_DIR="$HOME/.nvm"
[ -s "$NVM_DIR/nvm.sh" ] && \. "$NVM_DIR/nvm.sh"

print_info "Step 1: Setting up Node.js 22..."
nvm use 22 2>/dev/null || {
    print_warning "Node.js 22 not found, installing..."
    nvm install 22
    nvm use 22
}
print_status "Node.js $(node --version) ready"

print_info "Step 2: Setting up package manager..."
# Enable corepack for pnpm (auto-confirm)
export COREPACK_ENABLE_STRICT=0
echo "Y" | corepack enable 2>/dev/null || true

# Navigate to project
cd inbox-zero

print_info "Step 3: Installing dependencies..."
# Try pnpm install with auto-confirmation
echo "Y" | pnpm install 2>/dev/null || {
    print_warning "pnpm install failed, trying with force..."
    pnpm install --force
}

print_info "Step 4: Setting up database schema..."
cd apps/web
pnpm prisma generate
pnpm prisma db push

print_info "Step 5: Starting Inbox Zero..."
print_status "🎉 Launching your AI Email Assistant!"
print_info "Opening at: http://localhost:3000"
print_info "Press Ctrl+C to stop the server"

# Start the development server
pnpm dev
