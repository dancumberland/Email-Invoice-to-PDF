# Dream Pinball

A browser-based pinball game built with Phaser 3 and Matter.js physics.

## Project Status

**Phase**: Planning Complete, Ready to Build

## Features (Current - Vanilla JS Version)

- Custom table editor (draw walls, bumpers, targets, kickers, ramps)
- Physics-based ball movement with gravity and friction
- Dual flippers with keyboard controls
- Scoring system with combo multipliers
- Multiball mode
- Tilt detection (spam prevention)
- Sound effects (Web Audio API)
- Save/load table designs (localStorage)

## Features (Planned - Phaser Version)

- Matter.js physics with proper collision detection
- Modular architecture for graphics upgrades
- Enhanced visual effects (glow, particles, lighting)
- Multiple pre-designed tables
- High score persistence
- Mobile touch controls

## Tech Stack

| Component | Technology | Purpose |
|-----------|------------|---------|
| Framework | Phaser 3 | Game engine (physics + rendering) |
| Physics | Matter.js | Collision detection, rigid body dynamics |
| Rendering | Phaser WebGL | 2D rendering with WebGL acceleration |
| Build | Vite | Fast bundler and dev server |
| Deploy | Vercel | Static site hosting |

## Project Structure

```
Pax_Project_001/
├── docs/                    # Documentation
│   ├── README.md           # This file
│   ├── ARCHITECTURE.md     # System design
│   ├── PROJECT_BACKLOG.md  # Task tracking
│   └── research/           # Research findings
├── sessions/               # Session logs for AI continuity
├── src/                    # Source code (to be built)
│   ├── scenes/            # Phaser scenes
│   ├── objects/           # Game entities (Ball, Flipper, etc.)
│   ├── systems/           # Game systems (Physics, Audio, etc.)
│   ├── config/            # Configuration files
│   └── utils/             # Helper functions
├── public/                 # Static assets
│   └── assets/
│       ├── images/        # Sprites, backgrounds
│       ├── sounds/        # Audio files
│       └── tables/        # Table design JSON files
├── pinball02.html          # Original vanilla JS version (reference)
└── package.json            # Dependencies (to be created)
```

## Quick Start (After Build Phase)

```bash
# Install dependencies
npm install

# Start dev server
npm run dev

# Build for production
npm run build

# Deploy to Vercel
vercel --prod
```

## Controls

| Key | Action |
|-----|--------|
| Space | Launch ball |
| Left Arrow / A | Left flipper |
| Right Arrow / D | Right flipper |

## Development Notes

- See `docs/ARCHITECTURE.md` for system design
- See `docs/research/` for library research and decisions
- See `sessions/SESSIONS.md` for development history

## License

MIT
