# Project Backlog - Dream Pinball

**Last Updated**: December 4, 2025
**Session Context**: See [sessions/SESSIONS.md](../sessions/SESSIONS.md)

---

## NEXT (This or Next Session)

### ⭐ Learn Phaser Fundamentals
**Priority**: CRITICAL
**Estimated Effort**: 2-4 hours
**Status**: Ready to start

Before building, Dan and son should understand what Phaser is and how it works.

**Learning Resources**:

#### Videos (Recommended Order)

1. **"Phaser 3 Tutorial for Beginners"** - WornOffKeys
   - YouTube: https://www.youtube.com/watch?v=U0K8pLcFI8A
   - Duration: ~30 min
   - Why: Explains Phaser from scratch, very beginner-friendly

2. **"Make Your First Phaser 3 Game"** - Official Phaser Tutorial
   - YouTube: https://www.youtube.com/watch?v=5v5SnP7Dq_A
   - Duration: ~45 min
   - Why: Official tutorial, step-by-step simple game

3. **"Phaser 3 Physics with Matter.js"** - Game Dev Academy
   - YouTube: Search "Phaser 3 Matter.js tutorial"
   - Duration: ~20-30 min
   - Why: Shows exactly the physics system we'll use

4. **"Building a Breakout Game in Phaser 3"** - Ourcade
   - YouTube: https://www.youtube.com/watch?v=tpNZ3V31EWM
   - Duration: ~1 hour
   - Why: Breakout has similar physics to pinball (ball, paddles, bouncing)

#### Written Tutorials

- [Official "Making Your First Phaser 3 Game"](https://phaser.io/tutorials/making-your-first-phaser-3-game) - THE starting point
- [Phaser Examples Browser](https://phaser.io/examples) - Interactive code examples
- [Game Dev Academy Phaser Tutorials](https://gamedevacademy.org/category/phaser-tutorials/)

#### Quick Explanation for Son

**Phaser is like a LEGO set for making video games in a web browser.**

Instead of building every piece yourself:
- Drawing things on screen? Phaser does it.
- Making things bounce? Phaser does it.
- Keyboard controls? Phaser does it.
- Sound effects? Phaser does it.

You just tell Phaser WHAT you want, and it figures out HOW.

**Deliverables**:
- Watch at least videos 1 and 2
- Run the official "First Game" tutorial
- Understand: Scenes, Sprites, Physics, Input

**Dependencies**:
- None - can start immediately

---

## SOON (1-2 Weeks)

### Set Up Development Environment
**Priority**: HIGH
**Estimated Effort**: 1-2 hours
**Status**: Blocked by learning

Set up the project with Vite and Phaser.

**Deliverables**:
- `package.json` with Phaser and Vite
- Working "Hello Phaser" app
- Deployed to Vercel (even if empty)

**Dependencies**:
- ✅ Architecture planned
- ❌ Phaser fundamentals understood

---

### Port Ball Physics
**Priority**: HIGH
**Estimated Effort**: 2-3 hours
**Status**: Design phase

Convert the Ball class from vanilla JS to Phaser + Matter.js.

**Deliverables**:
- `src/objects/Ball.js` - Ball entity
- Ball spawns, falls with gravity, bounces off walls
- Ball trail effect

**Dependencies**:
- ❌ Dev environment set up

---

### Port Flipper Mechanics
**Priority**: HIGH
**Estimated Effort**: 3-4 hours
**Status**: Design phase

Convert flippers to Phaser + Matter.js with proper constraints.

**Deliverables**:
- `src/objects/Flipper.js` - Flipper entity
- Keyboard controls working
- Thick collision bodies (tunneling prevention)

**Dependencies**:
- ❌ Ball physics working

---

### Port Bumpers and Walls
**Priority**: HIGH
**Estimated Effort**: 2-3 hours
**Status**: Design phase

Convert static table elements.

**Deliverables**:
- `src/objects/Bumper.js`
- `src/objects/Wall.js`
- Default table layout

**Dependencies**:
- ❌ Flipper mechanics working

---

## BACKLOG (Future / Lower Priority)

### Add Targets, Kickers, Ramps
**Priority**: MEDIUM
**Estimated Effort**: 3-4 hours
**Status**: Design phase

Port special table elements.

**Deliverables**:
- `src/objects/Target.js`
- `src/objects/Kicker.js`
- `src/objects/Ramp.js`

---

### Implement Table Editor
**Priority**: MEDIUM
**Estimated Effort**: 4-6 hours
**Status**: Design phase

Recreate the drawing/editing mode from vanilla version.

**Deliverables**:
- `src/scenes/EditorScene.js`
- Draw walls, bumpers, targets
- Save/load table designs

---

### Add Sound System
**Priority**: LOW
**Estimated Effort**: 2 hours
**Status**: Design phase

Port Web Audio sounds to Phaser audio system.

**Deliverables**:
- `src/systems/AudioSystem.js`
- Flipper, bumper, launch sounds

---

### Add Visual Effects (Phase 2)
**Priority**: LOW
**Estimated Effort**: 4-6 hours
**Status**: Research complete

Add PixiJS filters for glow and bloom effects.

**Deliverables**:
- Glow on bumpers when hit
- Ball trail with bloom
- Neon-style lighting

**Dependencies**:
- ❌ Core game complete

---

### Mobile Touch Controls
**Priority**: LOW
**Estimated Effort**: 2-3 hours
**Status**: Not started

Add touch controls for mobile play.

**Deliverables**:
- Touch zones for left/right flippers
- Swipe to launch

---

## IDEAS (Exploratory / Unvalidated)

- Multiple table themes (space, underwater, retro)
- Online high score leaderboard
- Table sharing (export/import JSON)
- Achievements system
- Progressive Web App (PWA) for offline play
- VS mode (two players, split screen)
- VR/AR mode using WebXR (see pinball-xr research)

---

## BLOCKED

*No blocked items currently*

---

## Completed (This Session)

✅ Deep research on physics libraries (15+ evaluated)
✅ Deep research on graphics libraries (17+ evaluated)
✅ Architecture design with upgrade path
✅ Project folder structure created
✅ Documentation framework established
✅ Research documented for future reference
