# Graphics Libraries Research

**Date**: December 4, 2025
**Purpose**: Evaluate JavaScript graphics/rendering libraries for realistic pinball visuals

---

## Summary

After researching 17+ graphics libraries, **Phaser's built-in renderer** is recommended for Phase 1 (simplicity). **PixiJS** is the upgrade path for Phase 2 when we want glow effects, bloom, and advanced visuals.

---

## Libraries Evaluated

### Tier 1: Recommended

| Library | Stars | Type | Best For |
|---------|-------|------|----------|
| **PixiJS** | 46k | 2D WebGL | Maximum 2D performance + effects |
| **Phaser** | 39k | Game framework | Complete solution (our Phase 1) |
| **Babylon.js** | 25k | 3D engine | Best 2D sprite benchmarks |

### Tier 2: Good for Specific Use Cases

| Library | Stars | Type | Notes |
|---------|-------|------|-------|
| Three.js | 110k | 3D WebGL | Overkill for 2D, but best shaders |
| PlayCanvas | 14k | 3D engine | Visual editor, PBR |
| Konva.js | 14k | 2D Canvas | Good filters, React support |

### Tier 3: Not Recommended for Pinball

| Library | Reason |
|---------|--------|
| Fabric.js | Vector-focused, not game performance |
| p5.js | Creative coding, not optimized for games |
| Two.js | Slower in benchmarks |
| Stage.js | Possibly abandoned |

---

## Deep Dive: Top 3

### 1. PixiJS (Future Upgrade)

**GitHub**: https://github.com/pixijs/pixijs
**Stars**: 46,000+
**License**: MIT (FREE)

**Pros**:
- **Fastest 2D WebGL renderer** - proven in benchmarks
- Excellent filter system (glow, bloom, blur)
- Works with Matter.js
- Active development (v8.14.3 as of Nov 2025)
- WebGPU support

**Cons**:
- Requires manual physics integration
- More setup than Phaser

**Key Filters for Pinball**:
```javascript
import { GlowFilter } from '@pixi/filter-glow';
import { BloomFilter } from '@pixi/filter-bloom';

// Add glow to bumper
bumper.filters = [new GlowFilter({ color: 0xff0000, distance: 15 })];

// Add bloom for neon effect
ball.filters = [new BloomFilter({ blur: 2, brightness: 1.5 })];
```

**Resources**:
- [Official Site](https://pixijs.com/)
- [PixiJS Filters](https://github.com/pixijs/filters)
- [Matter.js + PixiJS Integration](https://github.com/celsowhite/matter-pixi)

---

### 2. Phaser 3 (Our Phase 1 Choice)

**GitHub**: https://github.com/phaserjs/phaser
**Stars**: 39,000+
**License**: MIT (FREE)

**Pros**:
- **All-in-one game framework**
- Built-in Matter.js physics
- Built-in WebGL rendering
- Massive tutorial library (700+ tutorials)
- Beginner-friendly

**Cons**:
- Less control over rendering pipeline
- Custom effects require WebGL pipeline knowledge

**Built-in Features**:
- Sprite rendering
- Particle systems
- Tweens and animations
- Camera effects
- Basic lighting (via plugins)

**Resources**:
- [Official Docs](https://docs.phaser.io/)
- [Making Your First Game](https://phaser.io/tutorials/making-your-first-phaser-3-game)
- [Phaser Examples](https://phaser.io/examples)

---

### 3. Babylon.js (Alternative for Maximum Realism)

**GitHub**: https://github.com/BabylonJS/Babylon.js
**Stars**: 25,000+
**License**: Apache 2.0 (FREE)

**Pros**:
- **Best 2D sprite performance in benchmarks** (beats all engines)
- Physically Based Rendering (PBR) for realistic materials
- Excellent particle systems
- Built-in physics options

**Cons**:
- 3D engine learning curve for 2D game
- Heavier than PixiJS
- More complexity than needed for Phase 1

**When to use**:
If we want 3D-quality metallic ball rendering while keeping 2D gameplay.

**Resources**:
- [Official Playground](https://playground.babylonjs.com/)
- [Documentation](https://doc.babylonjs.com/)

---

## Visual Effects for Pinball

### Phase 1 (Phaser Built-in)
- Basic particle effects on bumper hits
- Sprite scaling for "pop" animations
- Camera shake on big hits

### Phase 2 (PixiJS Filters)
| Effect | Filter | Use Case |
|--------|--------|----------|
| Glow | `@pixi/filter-glow` | Bumpers, targets, ball trail |
| Bloom | `@pixi/filter-bloom` | Neon lights, score popups |
| Blur | `@pixi/filter-blur` | Motion blur on fast ball |
| Displacement | `@pixi/filter-displacement` | Warped glass effect |

### Phase 3 (Custom Shaders)
- Metallic ball with environment reflection
- Dynamic lighting from bumpers
- Real-time shadows

---

## Performance Benchmarks

From [JS Game Rendering Benchmark](https://github.com/Shirajuki/js-game-rendering-benchmark):

| Library | 2D Sprites (10k) | Notes |
|---------|------------------|-------|
| **Babylon.js** | Best | Surprising winner for 2D |
| **PixiJS** | Excellent | Consistent top performer |
| **Phaser 3** | Very Good | Slightly behind PixiJS |
| Three.js | Good | Better for 3D |
| Two.js | Poor | Not recommended |

For pinball (hundreds of objects, not thousands), all top 3 perform identically well.

---

## Integration Patterns

### Phaser Only (Phase 1)
```javascript
// Everything in Phaser
const config = {
  type: Phaser.WEBGL,
  physics: { default: 'matter' },
  scene: [GameScene]
};
```

### Phaser + PixiJS Filters (Phase 2)
```javascript
// Add PixiJS filter to Phaser sprite
import { GlowFilter } from '@pixi/filter-glow';

// Access Phaser's internal PixiJS renderer
const phaserRenderer = this.game.renderer;
// Note: Requires custom WebGL pipeline - more complex
```

### PixiJS + Matter.js (Phase 3)
```javascript
// Separate physics and rendering
const engine = Matter.Engine.create();
const app = new PIXI.Application();

function gameLoop() {
  Matter.Engine.update(engine);
  // Sync PIXI sprites to Matter bodies
  ball.sprite.position = ball.body.position;
}
```

---

## Decision Matrix

| Factor | Phaser | PixiJS | Babylon.js |
|--------|--------|--------|------------|
| Ease of use | ⭐⭐⭐⭐⭐ | ⭐⭐⭐ | ⭐⭐ |
| Physics integration | ⭐⭐⭐⭐⭐ | ⭐⭐⭐ | ⭐⭐⭐⭐ |
| Visual effects | ⭐⭐⭐ | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| Documentation | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ | ⭐⭐⭐⭐ |
| Performance | ⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| Beginner-friendly | ⭐⭐⭐⭐⭐ | ⭐⭐⭐ | ⭐⭐ |

**Recommendation**:
- Phase 1: Phaser (fast development, working game)
- Phase 2: Add PixiJS filters for effects
- Phase 3 (optional): Full PixiJS migration for maximum control

---

## Cost Summary

| Library | License | Cost |
|---------|---------|------|
| PixiJS | MIT | **FREE** |
| Phaser 3 | MIT | **FREE** |
| Babylon.js | Apache 2.0 | **FREE** |
| Three.js | MIT | **FREE** |
| All filter packages | MIT | **FREE** |

**All recommended libraries are 100% free and open source.**
