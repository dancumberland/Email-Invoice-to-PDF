// ABOUTME: Table editor scene for creating and modifying pinball table layouts
// ABOUTME: Supports drawing walls, bumpers, targets, kickers, and ramps

import Phaser from 'phaser';

export default class EditorScene extends Phaser.Scene {
  constructor() {
    super('EditorScene');

    this.drawnWalls = [];
    this.bouncyWalls = [];
    this.targets = [];
    this.kickers = [];
    this.ramps = [];

    this.currentTool = 'draw';
    this.brushSize = 20;
    this.isDrawing = false;
  }

  create() {
    // Background
    this.add.rectangle(400, 500, 800, 1000, 0x2a2a2a);

    // Load any saved table data
    this.loadTableData();

    // Draw static table outline
    this.drawTableOutline();

    // Create tool graphics
    this.elementsGraphics = this.add.graphics();

    // Create UI
    this.createUI();

    // Set up input
    this.setupInput();
  }

  drawTableOutline() {
    const graphics = this.add.graphics();
    graphics.lineStyle(3, 0x888888);

    // Table boundary
    graphics.strokeRect(30, 42, 660, 916);

    // Launch lane divider
    graphics.strokeRect(700, 42, 70, 916);
  }

  createUI() {
    // Mode indicator
    this.modeText = this.add.text(400, 15, 'EDIT MODE', {
      fontSize: '24px',
      color: '#4CAF50',
      fontStyle: 'bold'
    });
    this.modeText.setOrigin(0.5, 0);

    // Tool buttons (left side)
    this.createToolButtons();

    // Instructions
    this.instructionsText = this.add.text(400, 970, 'Click and drag to draw. Press P to switch to Play Mode.', {
      fontSize: '14px',
      color: '#aaaaaa'
    });
    this.instructionsText.setOrigin(0.5);
  }

  createToolButtons() {
    const tools = [
      { key: 'draw', label: 'Wall', color: 0x888888, y: 80 },
      { key: 'bouncy', label: 'Bouncy', color: 0x4CAF50, y: 120 },
      { key: 'target', label: 'Target', color: 0xF44336, y: 160 },
      { key: 'kicker', label: 'Kicker', color: 0x4CAF50, y: 200 },
      { key: 'ramp', label: 'Ramp', color: 0x00BCD4, y: 240 },
      { key: 'erase', label: 'Erase', color: 0xFF5722, y: 280 }
    ];

    this.toolButtons = [];

    tools.forEach(tool => {
      // Button background
      const btn = this.add.rectangle(750, tool.y, 80, 30, tool.color)
        .setInteractive()
        .on('pointerdown', () => this.setTool(tool.key));

      // Button label
      const label = this.add.text(750, tool.y, tool.label, {
        fontSize: '12px',
        color: '#ffffff'
      });
      label.setOrigin(0.5);

      this.toolButtons.push({ key: tool.key, btn, label });
    });

    // Highlight current tool
    this.updateToolHighlight();

    // Brush size controls
    this.add.text(750, 330, 'Brush: ' + this.brushSize, {
      fontSize: '12px',
      color: '#ffffff'
    }).setOrigin(0.5);
    this.brushText = this.add.text(750, 330, 'Brush: ' + this.brushSize, {
      fontSize: '12px',
      color: '#ffffff'
    }).setOrigin(0.5);

    // Size buttons
    this.add.text(720, 360, '-', { fontSize: '20px', color: '#fff' })
      .setInteractive()
      .on('pointerdown', () => {
        this.brushSize = Math.max(10, this.brushSize - 5);
        this.brushText.setText('Brush: ' + this.brushSize);
      });

    this.add.text(770, 360, '+', { fontSize: '20px', color: '#fff' })
      .setInteractive()
      .on('pointerdown', () => {
        this.brushSize = Math.min(50, this.brushSize + 5);
        this.brushText.setText('Brush: ' + this.brushSize);
      });

    // Save/Load/Clear buttons
    this.add.text(750, 420, 'SAVE', {
      fontSize: '14px',
      color: '#4CAF50',
      backgroundColor: '#333',
      padding: { x: 10, y: 5 }
    }).setOrigin(0.5)
      .setInteractive()
      .on('pointerdown', () => this.saveTableData());

    this.add.text(750, 460, 'LOAD', {
      fontSize: '14px',
      color: '#2196F3',
      backgroundColor: '#333',
      padding: { x: 10, y: 5 }
    }).setOrigin(0.5)
      .setInteractive()
      .on('pointerdown', () => this.loadTableData());

    this.add.text(750, 500, 'CLEAR', {
      fontSize: '14px',
      color: '#f44336',
      backgroundColor: '#333',
      padding: { x: 10, y: 5 }
    }).setOrigin(0.5)
      .setInteractive()
      .on('pointerdown', () => this.clearAll());

    // Play mode button
    this.add.text(750, 560, 'PLAY', {
      fontSize: '18px',
      color: '#FFD700',
      backgroundColor: '#333',
      padding: { x: 15, y: 8 }
    }).setOrigin(0.5)
      .setInteractive()
      .on('pointerdown', () => this.switchToPlay());
  }

  setTool(tool) {
    this.currentTool = tool;
    this.updateToolHighlight();
  }

  updateToolHighlight() {
    this.toolButtons.forEach(item => {
      if (item.key === this.currentTool) {
        item.btn.setStrokeStyle(3, 0xFFFFFF);
      } else {
        item.btn.setStrokeStyle(0);
      }
    });
  }

  setupInput() {
    // Mouse/touch drawing
    this.input.on('pointerdown', (pointer) => {
      if (pointer.x < 690) {
        this.isDrawing = true;
        this.drawAtPosition(pointer.x, pointer.y);
      }
    });

    this.input.on('pointermove', (pointer) => {
      if (this.isDrawing && pointer.x < 690) {
        this.drawAtPosition(pointer.x, pointer.y);
      }
    });

    this.input.on('pointerup', () => {
      this.isDrawing = false;
    });

    // Keyboard
    this.input.keyboard.on('keydown-P', () => {
      this.switchToPlay();
    });
  }

  drawAtPosition(x, y) {
    const radius = this.brushSize / 2;
    const obj = { x, y, radius };

    switch (this.currentTool) {
      case 'draw':
        this.drawnWalls.push(obj);
        break;
      case 'bouncy':
        this.bouncyWalls.push(obj);
        break;
      case 'target':
        this.targets.push({ ...obj, hit: false, points: 100 });
        break;
      case 'kicker':
        this.kickers.push({ ...obj, active: false });
        break;
      case 'ramp':
        this.ramps.push(obj);
        break;
      case 'erase':
        this.eraseAtPosition(x, y);
        break;
    }

    this.redrawElements();
  }

  eraseAtPosition(x, y) {
    const eraseRadius = this.brushSize;

    const filterFn = (w) => {
      const dist = Math.sqrt((w.x - x) ** 2 + (w.y - y) ** 2);
      return dist > eraseRadius;
    };

    this.drawnWalls = this.drawnWalls.filter(filterFn);
    this.bouncyWalls = this.bouncyWalls.filter(filterFn);
    this.targets = this.targets.filter(filterFn);
    this.kickers = this.kickers.filter(filterFn);
    this.ramps = this.ramps.filter(filterFn);
  }

  redrawElements() {
    this.elementsGraphics.clear();

    // Draw walls (gray)
    this.elementsGraphics.fillStyle(0x888888);
    this.drawnWalls.forEach(w => {
      this.elementsGraphics.fillCircle(w.x, w.y, w.radius);
    });

    // Draw bouncy walls (green)
    this.elementsGraphics.fillStyle(0x4CAF50);
    this.bouncyWalls.forEach(w => {
      this.elementsGraphics.fillCircle(w.x, w.y, w.radius);
    });

    // Draw targets (red)
    this.elementsGraphics.fillStyle(0xF44336);
    this.targets.forEach(t => {
      this.elementsGraphics.fillCircle(t.x, t.y, t.radius);
    });

    // Draw kickers (green with arrow)
    this.elementsGraphics.fillStyle(0x4CAF50);
    this.kickers.forEach(k => {
      this.elementsGraphics.fillCircle(k.x, k.y, k.radius);
      this.elementsGraphics.fillStyle(0xFFFFFF);
      this.elementsGraphics.fillTriangle(
        k.x, k.y - 5,
        k.x - 4, k.y + 3,
        k.x + 4, k.y + 3
      );
      this.elementsGraphics.fillStyle(0x4CAF50);
    });

    // Draw ramps (cyan)
    this.elementsGraphics.fillStyle(0x00BCD4);
    this.ramps.forEach(r => {
      this.elementsGraphics.fillCircle(r.x, r.y, r.radius);
    });
  }

  saveTableData() {
    const data = {
      walls: this.drawnWalls,
      bouncy: this.bouncyWalls,
      targets: this.targets,
      kickers: this.kickers,
      ramps: this.ramps
    };

    localStorage.setItem('pinballTable', JSON.stringify(data));

    // Visual feedback
    const savedText = this.add.text(400, 500, 'SAVED!', {
      fontSize: '48px',
      color: '#4CAF50'
    });
    savedText.setOrigin(0.5);
    this.tweens.add({
      targets: savedText,
      alpha: 0,
      duration: 1000,
      onComplete: () => savedText.destroy()
    });
  }

  loadTableData() {
    const saved = localStorage.getItem('pinballTable');
    if (saved) {
      try {
        const data = JSON.parse(saved);
        this.drawnWalls = data.walls || [];
        this.bouncyWalls = data.bouncy || [];
        this.targets = data.targets || [];
        this.kickers = data.kickers || [];
        this.ramps = data.ramps || [];
        this.redrawElements();
      } catch (e) {
        console.error('Failed to load table data:', e);
      }
    }
  }

  clearAll() {
    this.drawnWalls = [];
    this.bouncyWalls = [];
    this.targets = [];
    this.kickers = [];
    this.ramps = [];
    this.redrawElements();
  }

  switchToPlay() {
    // Save current table data so GameScene can use it
    this.saveTableData();

    // Switch to GameScene
    this.scene.start('GameScene', {
      customTable: {
        walls: this.drawnWalls,
        bouncy: this.bouncyWalls,
        targets: this.targets,
        kickers: this.kickers,
        ramps: this.ramps
      }
    });
  }
}
