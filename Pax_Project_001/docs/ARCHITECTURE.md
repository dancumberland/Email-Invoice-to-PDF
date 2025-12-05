# Dream Pinball - Architecture

## Design Philosophy

**Modular Graphics Layer**: The architecture separates physics logic from rendering, allowing the graphics system to be upgraded (e.g., from Phaser's built-in renderer to PixiJS) without rewriting game logic.

**Upgrade Path**:
1. **Phase 1** (Current Plan): Phaser 3 + built-in Matter.js
2. **Phase 2** (Future): Add PixiJS filters for glow/bloom effects
3. **Phase 3** (Future): Migrate to PixiJS + Planck.js for maximum control

---

## System Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                         GAME LAYER                               │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────────────────┐  │
│  │ GameScene   │  │ TableEditor │  │ UI Scene                │  │
│  │ (play mode) │  │ (edit mode) │  │ (HUD, menus)            │  │
│  └──────┬──────┘  └──────┬──────┘  └────────────┬────────────┘  │
│         │                │                      │                │
│         └────────────────┼──────────────────────┘                │
│                          │                                       │
├──────────────────────────┼───────────────────────────────────────┤
│                    ENTITY LAYER                                  │
│  ┌─────────┐  ┌─────────┐  ┌─────────┐  ┌─────────┐  ┌───────┐  │
│  │  Ball   │  │ Flipper │  │ Bumper  │  │ Target  │  │ Wall  │  │
│  └────┬────┘  └────┬────┘  └────┬────┘  └────┬────┘  └───┬───┘  │
│       │            │            │            │           │       │
│       └────────────┴────────────┴────────────┴───────────┘       │
│                          │                                       │
├──────────────────────────┼───────────────────────────────────────┤
│                    SYSTEMS LAYER                                 │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐              │
│  │  Physics    │  │  Renderer   │  │   Audio     │              │
│  │  System     │  │  System     │  │   System    │              │
│  │ (Matter.js) │  │ (Phaser/    │  │ (Web Audio) │              │
│  │             │  │  PixiJS)    │  │             │              │
│  └─────────────┘  └─────────────┘  └─────────────┘              │
│                                                                  │
├──────────────────────────────────────────────────────────────────┤
│                    CONFIG LAYER                                  │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐              │
│  │  Physics    │  │   Table     │  │   Game      │              │
│  │  Config     │  │   Config    │  │   Config    │              │
│  │ (gravity,   │  │ (layouts,   │  │ (scoring,   │              │
│  │  bounce)    │  │  elements)  │  │  rules)     │              │
│  └─────────────┘  └─────────────┘  └─────────────┘              │
└──────────────────────────────────────────────────────────────────┘
```

---

## File Structure

```
src/
├── main.js                 # Phaser game initialization
├── scenes/
│   ├── BootScene.js       # Asset loading
│   ├── GameScene.js       # Main gameplay
│   ├── EditorScene.js     # Table editor
│   └── UIScene.js         # HUD overlay
├── objects/
│   ├── Ball.js            # Ball entity
│   ├── Flipper.js         # Flipper entity
│   ├── Bumper.js          # Bumper entity
│   ├── Target.js          # Drop target entity
│   ├── Kicker.js          # Kicker/slingshot entity
│   ├── Wall.js            # Wall segment entity
│   └── Ramp.js            # Ramp entity
├── systems/
│   ├── PhysicsSystem.js   # Physics abstraction layer
│   ├── RenderSystem.js    # Rendering abstraction layer (UPGRADE POINT)
│   ├── AudioSystem.js     # Sound management
│   ├── InputSystem.js     # Keyboard/touch handling
│   └── ScoreSystem.js     # Scoring and multipliers
├── config/
│   ├── physics.js         # Physics constants
│   ├── game.js            # Game rules and settings
│   └── tables/            # Table layout definitions
│       └── default.json   # Default table
└── utils/
    ├── collision.js       # Collision helpers
    └── math.js            # Math utilities
```

---

## Key Design Decisions

### 1. Renderer Abstraction (RenderSystem.js)

The `RenderSystem` wraps all drawing operations. This is the **key upgrade point** for future graphics enhancements.

```javascript
// ABOUTME: Abstracts rendering operations for future graphics upgrades
// ABOUTME: Swap this system to migrate from Phaser to PixiJS

class RenderSystem {
  constructor(scene) {
    this.scene = scene;
  }

  // Entity rendering methods - override these for PixiJS migration
  drawBall(ball) { /* Phaser implementation */ }
  drawFlipper(flipper) { /* Phaser implementation */ }
  drawBumper(bumper) { /* Phaser implementation */ }

  // Effect methods - add PixiJS filters here later
  addGlowEffect(entity) { /* No-op for now, add PixiJS later */ }
  addParticleEffect(x, y, type) { /* Basic Phaser particles */ }
}
```

### 2. Physics Abstraction (PhysicsSystem.js)

Wraps Matter.js operations. If we need to migrate to Planck.js for better collision detection, we change this one file.

```javascript
// ABOUTME: Abstracts physics operations for potential engine swaps
// ABOUTME: Swap this system to migrate from Matter.js to Planck.js

class PhysicsSystem {
  constructor(scene) {
    this.scene = scene;
    this.engine = scene.matter.world;
  }

  createBall(x, y, radius) { /* Matter.js body */ }
  createFlipper(x, y, config) { /* Matter.js constraint */ }
  applyForce(body, force) { /* Matter.js force */ }

  // Collision callback registration
  onCollision(callback) { /* Matter.js events */ }
}
```

### 3. Entity Pattern

Each game object is a self-contained entity with:
- **Physics body** (managed by PhysicsSystem)
- **Visual representation** (managed by RenderSystem)
- **Game logic** (scoring, behavior)

```javascript
// ABOUTME: Base entity class for all game objects
// ABOUTME: Separates physics body from visual sprite

class Entity {
  constructor(scene, x, y) {
    this.scene = scene;
    this.physics = scene.systems.physics;
    this.renderer = scene.systems.renderer;
    this.body = null;   // Physics body
    this.sprite = null; // Visual representation
  }

  update(delta) { /* Override in subclass */ }
  destroy() { /* Cleanup physics and visuals */ }
}
```

### 4. Table Data Format

Tables are defined as JSON, separate from code:

```json
{
  "name": "Classic Table",
  "dimensions": { "width": 800, "height": 1000 },
  "elements": {
    "walls": [
      { "type": "line", "start": [30, 50], "end": [30, 950] },
      { "type": "circle", "x": 100, "y": 200, "radius": 20 }
    ],
    "bumpers": [
      { "x": 150, "y": 200, "width": 80, "height": 60, "points": 100 }
    ],
    "flippers": {
      "left": { "x": 300, "y": 920 },
      "right": { "x": 500, "y": 920 }
    },
    "targets": [],
    "kickers": [],
    "ramps": []
  },
  "physics": {
    "gravity": 0.25,
    "friction": 0.988,
    "launchPower": 28
  }
}
```

---

## Upgrade Path Details

### Phase 1: Basic Phaser Implementation
- Use Phaser's built-in Matter.js plugin
- Use Phaser's built-in WebGL renderer
- Simple sprites and shapes
- **Goal**: Working game with good physics

### Phase 2: Enhanced Visuals (PixiJS Filters)
- Add `@pixi/filter-*` packages
- Implement glow effects on bumpers and ball
- Add bloom post-processing
- Particle effects on hits
- **Changes**: Only `RenderSystem.js` modifications

### Phase 3: Maximum Realism (Optional)
- Migrate physics to Planck.js (better CCD)
- Full PixiJS rendering
- Custom shaders for metallic ball
- Dynamic lighting
- **Changes**: Replace `PhysicsSystem.js` and `RenderSystem.js`

---

## Collision Detection Strategy

### The Tunneling Problem

Fast balls can pass through thin flippers. Our solution:

1. **Thick collision bodies** on flippers (invisible, larger than visual)
2. **Multiple collision checks per frame** for fast-moving objects
3. **Fallback**: Migrate to Planck.js if issues persist

```javascript
// Flipper collision body is thicker than visual
createFlipper(x, y, isLeft) {
  const visualWidth = 15;
  const collisionWidth = 30; // 2x visual for safety

  // Visual sprite uses visualWidth
  // Physics body uses collisionWidth
}
```

---

## State Management

### Game State
```javascript
const gameState = {
  score: 0,
  ballsLeft: 3,
  multiplier: 1,
  tilted: false,
  mode: 'play' | 'edit',
  currentTable: 'default'
};
```

### Table State (Editor)
```javascript
const tableState = {
  walls: [],
  bumpers: [],
  targets: [],
  kickers: [],
  ramps: [],
  modified: false
};
```

---

## Dependencies

```json
{
  "dependencies": {
    "phaser": "^3.80.0"
  },
  "devDependencies": {
    "vite": "^5.0.0"
  }
}
```

### Future Dependencies (Phase 2+)
```json
{
  "@pixi/filter-glow": "^5.0.0",
  "@pixi/filter-bloom": "^5.0.0",
  "planck": "^1.0.0"
}
```

---

## Build & Deploy

### Development
```bash
npm run dev    # Vite dev server at localhost:5173
```

### Production
```bash
npm run build  # Outputs to dist/
vercel --prod  # Deploys to Vercel
```

### Vercel Configuration
- Framework: Other
- Output Directory: dist
- Build Command: npm run build

---

## Testing Strategy

1. **Manual testing** for physics feel
2. **Visual debugging** with Matter.js debug renderer
3. **Performance monitoring** via browser devtools
4. **Collision edge cases** with high-speed ball launches

---

## References

- [Phaser 3 Documentation](https://docs.phaser.io/)
- [Matter.js Documentation](https://brm.io/matter-js/docs/)
- [Coder's Block Pinball Tutorial](https://codersblock.com/blog/javascript-physics-with-matter-js/)
- See `docs/research/` for library evaluation details
