# Physics Libraries Research

**Date**: December 4, 2025
**Purpose**: Evaluate JavaScript physics engines for realistic pinball game

---

## Summary

After researching 15+ physics engines, **Matter.js** (via Phaser) is recommended for Phase 1 due to its ease of use and integration. **Planck.js** is the backup if collision tunneling becomes problematic.

---

## Libraries Evaluated

### Tier 1: Recommended

| Library | Stars | Status | CCD Support | Best For |
|---------|-------|--------|-------------|----------|
| **Matter.js** | 16k+ | Active | No (workaround needed) | Beginners, Phaser integration |
| **Planck.js** | 4.7k | Active | Yes | When Matter.js tunneling fails |
| **Box2D (WASM)** | Mature | Active | Yes | Industry-standard reliability |

### Tier 2: Good Alternatives

| Library | Stars | Status | Notes |
|---------|-------|--------|-------|
| p2.js | 2.8k | Less active | Lightweight, mobile-friendly |
| Rapier | 3k+ | Active | Rust-based WASM, very fast |

### Tier 3: Not Recommended for 2D Pinball

| Library | Reason |
|---------|--------|
| Cannon.js | 3D-focused, overkill |
| Ammo.js | Too heavy for 2D |
| Oimo.js | 3D-focused |
| Physics.js | Abandoned |
| Verlet-js | Wrong use case (soft body) |

---

## Deep Dive: Top 3

### 1. Matter.js

**GitHub**: https://github.com/liabru/matter-js
**Stars**: 16,000+
**License**: MIT (FREE)

**Pros**:
- Built into Phaser 3 (zero integration work)
- Excellent documentation
- Large community
- Easy to learn API
- Works in browser without build step

**Cons**:
- **No continuous collision detection (CCD)** - fast balls can tunnel through thin flippers
- Less precise than Box2D

**Tunneling Workaround**:
```javascript
// Make flipper collision body thicker than visual
const visualWidth = 15;
const collisionWidth = 30; // Invisible but catches fast balls
```

**Resources**:
- [Official Docs](https://brm.io/matter-js/docs/)
- [Coder's Block Pinball Tutorial](https://codersblock.com/blog/javascript-physics-with-matter-js/) - MUST READ
- [Phaser Matter.js Examples](https://phaser.io/examples/v3/category/physics/matterjs)

---

### 2. Planck.js

**GitHub**: https://github.com/piqnt/planck.js
**Stars**: 4,700+
**License**: MIT (FREE)

**Pros**:
- Box2D algorithms rewritten in JavaScript
- **Has CCD** - solves tunneling problem
- Deterministic (same input = same output)
- Good for competitive/replay features

**Cons**:
- Requires manual integration with graphics
- Steeper learning curve than Matter.js
- Smaller community

**When to use**:
If Matter.js tunneling workarounds fail, migrate physics to Planck.js.

**Resources**:
- [Official Site](https://piqnt.com/planck.js/)
- [GitHub Examples](https://github.com/piqnt/planck.js/tree/master/example)

---

### 3. Box2D (box2d-wasm)

**GitHub**: https://github.com/nicksherron/box2d-wasm
**License**: MIT (FREE)

**Pros**:
- Industry standard (used in Angry Birds, etc.)
- Most mature physics engine
- Excellent collision detection
- Deterministic

**Cons**:
- WASM adds complexity
- Overkill for browser game
- Harder to debug

**When to use**:
Only if building a professional/commercial game requiring absolute physics reliability.

---

## The Tunneling Problem Explained

**What is it?**
When physics runs at 60fps and a ball moves 20 pixels per frame, it can "teleport" through a 15-pixel-wide flipper.

```
Frame 1: Ball at x=100 (before flipper)
Frame 2: Ball at x=120 (past flipper!)
         Flipper is at x=110, width=15
         Ball never "touched" flipper in any frame
```

**Solutions**:

1. **Thick collision bodies** (Matter.js workaround)
   - Visual flipper: 15px wide
   - Collision body: 40px wide (invisible)

2. **Multiple sub-steps** per frame
   - Run physics 2-4x per render frame
   - Performance cost

3. **CCD engines** (Planck.js, Box2D)
   - Engine traces ball path between frames
   - Detects intersection even if not at frame boundary

---

## Decision Matrix

| Factor | Matter.js | Planck.js | Box2D |
|--------|-----------|-----------|-------|
| Ease of use | ⭐⭐⭐⭐⭐ | ⭐⭐⭐ | ⭐⭐ |
| Phaser integration | ⭐⭐⭐⭐⭐ | ⭐⭐ | ⭐⭐ |
| Collision accuracy | ⭐⭐⭐ | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| Documentation | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ | ⭐⭐⭐ |
| Community size | ⭐⭐⭐⭐⭐ | ⭐⭐⭐ | ⭐⭐⭐ |
| Performance | ⭐⭐⭐⭐ | ⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ |

**Recommendation**: Start with Matter.js (via Phaser). If tunneling is unacceptable after implementing thick collision bodies, migrate PhysicsSystem to Planck.js.

---

## Open-Source Pinball Projects Studied

| Project | Stack | Stars | Key Learning |
|---------|-------|-------|--------------|
| [fishshiz/pinball-wizard](https://github.com/fishshiz/pinball-wizard) | Matter.js + Canvas | 7 | Clean architecture |
| [ag-game/phaser-pinball](https://github.com/ag-game/phaser-pinball) | Phaser + Matter.js | 1 | Phaser patterns |
| [vpdb/vpx-js](https://github.com/vpdb/vpx-js) | Three.js + Custom | 59 | Pro TypeScript |
| [h4k1m0u/pinball](https://github.com/h4k1m0u/pinball) | Planck.js + p5.js | 1 | CCD benefits |
| [lrusso/Pinball](https://github.com/lrusso/Pinball) | Phaser v2 | 10 | PWA approach |

---

## Cost Summary

| Library | License | Cost |
|---------|---------|------|
| Matter.js | MIT | **FREE** |
| Planck.js | MIT | **FREE** |
| Box2D | MIT | **FREE** |
| p2.js | MIT | **FREE** |
| Rapier | Apache 2.0 | **FREE** |

**All recommended libraries are 100% free and open source.**
