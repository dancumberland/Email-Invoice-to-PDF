// ABOUTME: Main gameplay scene with pinball table, ball, and flippers
// ABOUTME: Handles physics simulation and player input

import Phaser from 'phaser';

export default class GameScene extends Phaser.Scene {
  constructor() {
    super('GameScene');

    this.ball = null;
    this.leftFlipper = null;
    this.rightFlipper = null;
    this.walls = [];
    this.bumpers = [];

    this.score = 0;
    this.ballsLeft = 3;
    this.launched = false;

    this.targets = [];
    this.kickers = [];
    this.ramps = [];
    this.multiplier = 1;
    this.lastHitTime = 0;
    this.comboTimer = 0;

    this.customTable = null;
  }

  init(data) {
    // Receive custom table data from editor
    if (data && data.customTable) {
      this.customTable = data.customTable;
    } else {
      this.customTable = null;
    }

    // Reset state
    this.score = 0;
    this.ballsLeft = 3;
    this.multiplier = 1;
    this.targets = [];
    this.kickers = [];
    this.ramps = [];
  }

  create() {
    // Create table walls
    this.createWalls();

    // Create bumpers
    this.createBumpers();

    // Create targets
    this.createTargets();

    // Create kickers
    this.createKickers();

    // Create ramps
    this.createRamps();

    // Create flippers
    this.createFlippers();

    // Create ball
    this.createBall();

    // Create UI
    this.createUI();

    // Set up input
    this.setupInput();

    // Collision handling
    this.setupCollisions();
  }

  createWalls() {
    // Wall positions matching EditorScene strokeRect(30, 42, 660, 916) and strokeRect(700, 42, 70, 916)
    // Main playfield: x=30-690, y=42-958
    // Launch lane: x=700-770, y=42-958

    // Left wall - thin wall at x=30
    const leftWall = this.matter.add.rectangle(30, 500, 6, 916, {
      isStatic: true,
      label: 'wall'
    });

    // Right wall of playfield - at x=690
    const rightWall = this.matter.add.rectangle(690, 500, 6, 916, {
      isStatic: true,
      label: 'wall'
    });

    // Launch lane left wall - at x=700
    const launchLeftWall = this.matter.add.rectangle(700, 500, 6, 916, {
      isStatic: true,
      label: 'wall'
    });

    // Launch lane right wall - at x=770
    const launchRightWall = this.matter.add.rectangle(770, 500, 6, 916, {
      isStatic: true,
      label: 'wall'
    });

    // Top wall - at y=42
    const topWall = this.matter.add.rectangle(400, 42, 740, 6, {
      isStatic: true,
      label: 'wall'
    });

    // Bottom walls with gap for drain - at y=958
    const bottomLeft = this.matter.add.rectangle(150, 958, 240, 6, {
      isStatic: true,
      label: 'wall'
    });
    const bottomRight = this.matter.add.rectangle(550, 958, 280, 6, {
      isStatic: true,
      label: 'wall'
    });

    // Flipper area slingshot walls
    const leftFlipperWall = this.matter.add.rectangle(180, 870, 6, 100, {
      isStatic: true,
      angle: 0.5,
      label: 'wall'
    });
    const rightFlipperWall = this.matter.add.rectangle(520, 870, 6, 100, {
      isStatic: true,
      angle: -0.5,
      label: 'wall'
    });

    this.walls = [leftWall, rightWall, launchLeftWall, launchRightWall, topWall, bottomLeft, bottomRight, leftFlipperWall, rightFlipperWall];

    // Create custom drawn walls from editor
    this.createCustomWalls();

    // Draw wall visuals
    this.drawWalls();
  }

  createCustomWalls() {
    if (!this.customTable) return;

    // Create physics bodies for drawn walls
    if (this.customTable.walls && this.customTable.walls.length > 0) {
      this.customTable.walls.forEach(data => {
        const wall = this.matter.add.circle(data.x, data.y, data.radius, {
          isStatic: true,
          label: 'wall'
        });
        this.walls.push(wall);
      });
    }

    // Create physics bodies for bouncy walls
    if (this.customTable.bouncy && this.customTable.bouncy.length > 0) {
      this.customTable.bouncy.forEach(data => {
        const bouncyWall = this.matter.add.circle(data.x, data.y, data.radius, {
          isStatic: true,
          restitution: 1.3,
          label: 'bumper'
        });
        this.bumpers.push(bouncyWall);
      });
    }
  }

  drawWalls() {
    const graphics = this.add.graphics();

    // Match EditorScene exactly - table boundary and launch lane
    graphics.lineStyle(3, 0x888888);
    graphics.strokeRect(30, 42, 660, 916);
    graphics.strokeRect(700, 42, 70, 916);

    // Draw custom walls from editor
    if (this.customTable) {
      // Regular walls (gray)
      if (this.customTable.walls) {
        graphics.fillStyle(0x888888);
        this.customTable.walls.forEach(w => {
          graphics.fillCircle(w.x, w.y, w.radius);
        });
      }

      // Bouncy walls (green)
      if (this.customTable.bouncy) {
        graphics.fillStyle(0x4CAF50);
        this.customTable.bouncy.forEach(w => {
          graphics.fillCircle(w.x, w.y, w.radius);
        });
      }
    }
  }

  createBumpers() {
    const bumperData = [
      { x: 190, y: 230, w: 80, h: 60, points: 100 },
      { x: 330, y: 275, w: 60, h: 50, points: 150 },
      { x: 560, y: 235, w: 80, h: 70, points: 100 },
      { x: 175, y: 430, w: 70, h: 60, points: 200 },
      { x: 395, y: 475, w: 90, h: 50, points: 150 },
      { x: 585, y: 455, w: 70, h: 80, points: 200 }
    ];

    const graphics = this.add.graphics();

    bumperData.forEach(data => {
      const bumper = this.matter.add.rectangle(
        data.x + data.w / 2,
        data.y + data.h / 2,
        data.w,
        data.h,
        {
          isStatic: true,
          restitution: 1.3,
          label: 'bumper',
          points: data.points
        }
      );

      // Draw bumper visual
      graphics.fillStyle(0x4CAF50);
      graphics.fillRect(data.x, data.y, data.w, data.h);
      graphics.lineStyle(3, 0x6C6C6C);
      graphics.strokeRect(data.x, data.y, data.w, data.h);

      this.bumpers.push(bumper);
    });
  }

  createTargets() {
    // Use custom table data if available, otherwise use defaults
    const targetData = (this.customTable && this.customTable.targets && this.customTable.targets.length > 0)
      ? this.customTable.targets
      : [
          { x: 100, y: 350, radius: 15, points: 100 },
          { x: 650, y: 320, radius: 15, points: 100 },
          { x: 250, y: 550, radius: 15, points: 150 },
          { x: 550, y: 580, radius: 15, points: 150 }
        ];

    const graphics = this.add.graphics();

    targetData.forEach(data => {
      const target = this.matter.add.circle(data.x, data.y, data.radius, {
        isStatic: true,
        isSensor: true,
        label: 'target'
      });
      target.points = data.points || 100;
      target.hit = false;
      target.hitTime = 0;
      target.graphics = graphics;
      target.baseX = data.x;
      target.baseY = data.y;
      target.radius = data.radius;

      this.targets.push(target);
    });
  }

  createKickers() {
    // Use custom table data if available, otherwise use defaults
    const kickerData = (this.customTable && this.customTable.kickers && this.customTable.kickers.length > 0)
      ? this.customTable.kickers
      : [
          { x: 80, y: 700, radius: 20 },
          { x: 620, y: 680, radius: 20 }
        ];

    kickerData.forEach(data => {
      const kicker = this.matter.add.circle(data.x, data.y, data.radius, {
        isStatic: true,
        isSensor: true,
        label: 'kicker'
      });
      kicker.active = false;
      kicker.baseX = data.x;
      kicker.baseY = data.y;
      kicker.radius = data.radius;

      this.kickers.push(kicker);
    });
  }

  createRamps() {
    // Use custom table data if available, otherwise use defaults
    const rampData = (this.customTable && this.customTable.ramps && this.customTable.ramps.length > 0)
      ? this.customTable.ramps
      : [
          { x: 150, y: 600, radius: 25 },
          { x: 550, y: 620, radius: 25 }
        ];

    rampData.forEach(data => {
      const ramp = this.matter.add.circle(data.x, data.y, data.radius, {
        isStatic: true,
        isSensor: true,
        label: 'ramp'
      });
      ramp.baseX = data.x;
      ramp.baseY = data.y;
      ramp.radius = data.radius;

      this.ramps.push(ramp);
    });
  }

  createFlippers() {
    const flipperLength = 100;
    const flipperWidth = 15;
    const MatterConstraint = Phaser.Physics.Matter.Matter.Constraint;

    // Flipper Y position - above the drain gap at y=958, accounting for slingshots
    const flipperY = 920;

    // Left flipper - positioned to meet the left slingshot wall
    const leftPivotX = 250;
    const leftPivotY = flipperY;

    this.leftFlipperBody = this.matter.add.rectangle(
      leftPivotX + flipperLength / 2,
      leftPivotY,
      flipperLength,
      flipperWidth,
      {
        label: 'flipper'
      }
    );

    // Pin constraint for left flipper pivot (world-pinned)
    const leftConstraint = MatterConstraint.create({
      bodyA: this.leftFlipperBody,
      pointA: { x: -flipperLength / 2, y: 0 },
      pointB: { x: leftPivotX, y: leftPivotY },
      stiffness: 1,
      length: 0
    });
    this.matter.world.add(leftConstraint);

    // Right flipper - positioned to meet the right slingshot wall
    const rightPivotX = 450;
    const rightPivotY = flipperY;

    this.rightFlipperBody = this.matter.add.rectangle(
      rightPivotX - flipperLength / 2,
      rightPivotY,
      flipperLength,
      flipperWidth,
      {
        label: 'flipper'
      }
    );

    // Pin constraint for right flipper pivot (world-pinned)
    const rightConstraint = MatterConstraint.create({
      bodyA: this.rightFlipperBody,
      pointA: { x: flipperLength / 2, y: 0 },
      pointB: { x: rightPivotX, y: rightPivotY },
      stiffness: 1,
      length: 0
    });
    this.matter.world.add(rightConstraint);

    // Set initial angles
    this.matter.body.setAngle(this.leftFlipperBody, 0.4);
    this.matter.body.setAngle(this.rightFlipperBody, -0.4);

    // Flipper graphics will be updated in update loop
    this.leftFlipperGraphics = this.add.graphics();
    this.rightFlipperGraphics = this.add.graphics();
  }

  createBall() {
    // Ball starts in launch lane (centered between x=700 and x=770 = 735)
    this.ball = this.matter.add.circle(735, 850, 10, {
      restitution: 0.6,
      friction: 0.001,
      frictionAir: 0.01,
      label: 'ball'
    });

    this.ballGraphics = this.add.graphics();
    this.ballTrail = [];

    // Prevent ball from moving until launched
    this.matter.body.setStatic(this.ball, true);
    this.launched = false;

    // Launch text
    this.launchText = this.add.text(735, 800, 'PRESS\nSPACE', {
      fontSize: '14px',
      color: '#ffffff',
      align: 'center'
    });
    this.launchText.setOrigin(0.5);
  }

  createUI() {
    // Score display
    this.scoreText = this.add.text(100, 10, 'Score: 0', {
      fontSize: '24px',
      color: '#ffffff'
    });

    // Balls display
    this.ballsText = this.add.text(300, 10, 'Balls: 3', {
      fontSize: '24px',
      color: '#ffffff'
    });

    // Multiplier display
    this.multiplierText = this.add.text(500, 10, 'x1', {
      fontSize: '28px',
      color: '#FFD700'
    });

    // Edit button
    this.add.text(750, 15, 'EDIT', {
      fontSize: '18px',
      color: '#4CAF50',
      backgroundColor: '#333',
      padding: { x: 10, y: 5 }
    })
      .setInteractive()
      .on('pointerdown', () => this.switchToEditor());
  }

  switchToEditor() {
    this.scene.start('EditorScene');
  }

  setupInput() {
    this.cursors = this.input.keyboard.createCursorKeys();
    this.keyA = this.input.keyboard.addKey(Phaser.Input.Keyboard.KeyCodes.A);
    this.keyD = this.input.keyboard.addKey(Phaser.Input.Keyboard.KeyCodes.D);
    this.spaceKey = this.input.keyboard.addKey(Phaser.Input.Keyboard.KeyCodes.SPACE);
  }

  setupCollisions() {
    this.matter.world.on('collisionstart', (event) => {
      event.pairs.forEach(pair => {
        const labels = [pair.bodyA.label, pair.bodyB.label];

        if (labels.includes('ball') && labels.includes('bumper')) {
          const bumper = pair.bodyA.label === 'bumper' ? pair.bodyA : pair.bodyB;
          this.hitBumper(bumper);
        }

        if (labels.includes('ball') && labels.includes('target')) {
          const target = pair.bodyA.label === 'target' ? pair.bodyA : pair.bodyB;
          this.hitTarget(target);
        }

        if (labels.includes('ball') && labels.includes('kicker')) {
          const kicker = pair.bodyA.label === 'kicker' ? pair.bodyA : pair.bodyB;
          this.hitKicker(kicker);
        }

        if (labels.includes('ball') && labels.includes('ramp')) {
          this.hitRamp();
        }
      });
    });
  }

  hitBumper(bumper) {
    const points = bumper.points || 100;
    this.addScore(points);
    this.updateCombo();
  }

  hitTarget(target) {
    if (target.hit) return;

    target.hit = true;
    target.hitTime = Date.now();
    this.addScore(target.points || 100);
    this.updateCombo();

    // Bounce ball away
    const angle = Math.atan2(
      this.ball.position.y - target.baseY,
      this.ball.position.x - target.baseX
    );
    this.matter.body.setVelocity(this.ball, {
      x: Math.cos(angle) * 8,
      y: Math.sin(angle) * 8
    });

    // Reset after 3 seconds
    this.time.delayedCall(3000, () => {
      target.hit = false;
    });
  }

  hitKicker(kicker) {
    if (kicker.active) return;

    kicker.active = true;
    this.addScore(75);

    // Delay then launch ball upward
    this.time.delayedCall(200, () => {
      if (this.ball) {
        this.matter.body.setVelocity(this.ball, {
          x: (Math.random() - 0.5) * 10,
          y: -20
        });
      }
      kicker.active = false;
    });
  }

  hitRamp() {
    // Speed boost
    if (this.ball) {
      const vel = this.ball.velocity;
      this.matter.body.setVelocity(this.ball, {
        x: vel.x * 1.1,
        y: vel.y * 1.1
      });
    }
  }

  addScore(points) {
    this.score += points * this.multiplier;
    this.scoreText.setText('Score: ' + this.score);
  }

  updateCombo() {
    const now = Date.now();
    if (now - this.lastHitTime < 1000) {
      this.multiplier = Math.min(this.multiplier + 1, 10);
      this.comboTimer = 60;
    } else {
      this.multiplier = 1;
      this.comboTimer = 60;
    }
    this.lastHitTime = now;
    this.multiplierText.setText('x' + this.multiplier);
  }

  update() {
    // Handle flipper controls
    this.updateFlippers();

    // Handle ball launch
    this.handleLaunch();

    // Draw ball
    this.drawBall();

    // Draw flippers
    this.drawFlippers();

    // Draw targets, kickers, ramps
    this.drawTargets();
    this.drawKickers();
    this.drawRamps();

    // Update combo timer
    this.updateComboTimer();

    // Check if ball is lost
    this.checkBallLost();
  }

  updateFlippers() {
    const flipperForce = 0.15;
    const restAngleLeft = 0.4;
    const restAngleRight = -0.4;
    const activeAngleLeft = -0.6;
    const activeAngleRight = 0.6;

    // Left flipper
    if (this.cursors.left.isDown || this.keyA.isDown) {
      const targetAngle = activeAngleLeft;
      const currentAngle = this.leftFlipperBody.angle;
      const angleDiff = targetAngle - currentAngle;
      this.matter.body.setAngularVelocity(this.leftFlipperBody, angleDiff * 0.5);
    } else {
      const targetAngle = restAngleLeft;
      const currentAngle = this.leftFlipperBody.angle;
      const angleDiff = targetAngle - currentAngle;
      this.matter.body.setAngularVelocity(this.leftFlipperBody, angleDiff * 0.3);
    }

    // Right flipper
    if (this.cursors.right.isDown || this.keyD.isDown) {
      const targetAngle = activeAngleRight;
      const currentAngle = this.rightFlipperBody.angle;
      const angleDiff = targetAngle - currentAngle;
      this.matter.body.setAngularVelocity(this.rightFlipperBody, angleDiff * 0.5);
    } else {
      const targetAngle = restAngleRight;
      const currentAngle = this.rightFlipperBody.angle;
      const angleDiff = targetAngle - currentAngle;
      this.matter.body.setAngularVelocity(this.rightFlipperBody, angleDiff * 0.3);
    }
  }

  handleLaunch() {
    if (!this.launched && Phaser.Input.Keyboard.JustDown(this.spaceKey)) {
      this.matter.body.setStatic(this.ball, false);
      this.matter.body.setVelocity(this.ball, { x: 0, y: -25 });
      this.launched = true;
      this.launchText.setVisible(false);
    }
  }

  drawBall() {
    this.ballGraphics.clear();

    if (!this.ball) return;

    const x = this.ball.position.x;
    const y = this.ball.position.y;

    // Draw trail
    this.ballTrail.push({ x, y });
    if (this.ballTrail.length > 15) {
      this.ballTrail.shift();
    }

    this.ballGraphics.lineStyle(3, 0xffffff, 0.3);
    this.ballGraphics.beginPath();
    this.ballTrail.forEach((point, i) => {
      if (i === 0) {
        this.ballGraphics.moveTo(point.x, point.y);
      } else {
        this.ballGraphics.lineTo(point.x, point.y);
      }
    });
    this.ballGraphics.strokePath();

    // Draw ball
    this.ballGraphics.fillStyle(0xffffff);
    this.ballGraphics.fillCircle(x, y, 10);

    // Glow effect
    this.ballGraphics.fillStyle(0xffffff, 0.3);
    this.ballGraphics.fillCircle(x, y, 15);
  }

  drawFlippers() {
    // Left flipper
    this.leftFlipperGraphics.clear();
    this.leftFlipperGraphics.fillStyle(0xbbbbbb);

    const leftPos = this.leftFlipperBody.position;
    const leftAngle = this.leftFlipperBody.angle;

    this.leftFlipperGraphics.save();
    this.leftFlipperGraphics.translateCanvas(leftPos.x, leftPos.y);
    this.leftFlipperGraphics.rotateCanvas(leftAngle);
    this.leftFlipperGraphics.fillRoundedRect(-50, -7.5, 100, 15, 7.5);
    this.leftFlipperGraphics.restore();

    // Pivot point
    this.leftFlipperGraphics.fillStyle(0x666666);
    this.leftFlipperGraphics.fillCircle(280, 900, 10);

    // Right flipper
    this.rightFlipperGraphics.clear();
    this.rightFlipperGraphics.fillStyle(0xbbbbbb);

    const rightPos = this.rightFlipperBody.position;
    const rightAngle = this.rightFlipperBody.angle;

    this.rightFlipperGraphics.save();
    this.rightFlipperGraphics.translateCanvas(rightPos.x, rightPos.y);
    this.rightFlipperGraphics.rotateCanvas(rightAngle);
    this.rightFlipperGraphics.fillRoundedRect(-50, -7.5, 100, 15, 7.5);
    this.rightFlipperGraphics.restore();

    // Pivot point
    this.rightFlipperGraphics.fillStyle(0x666666);
    this.rightFlipperGraphics.fillCircle(520, 900, 10);
  }

  drawTargets() {
    if (!this.targetGraphics) {
      this.targetGraphics = this.add.graphics();
    }
    this.targetGraphics.clear();

    this.targets.forEach(target => {
      // Red when active, yellow when hit
      const color = target.hit ? 0xFFFF00 : 0xF44336;
      this.targetGraphics.fillStyle(color);
      this.targetGraphics.fillCircle(target.baseX, target.baseY, target.radius);

      if (target.hit) {
        this.targetGraphics.lineStyle(4, 0xFFFF00);
        this.targetGraphics.strokeCircle(target.baseX, target.baseY, target.radius);
      }
    });
  }

  drawKickers() {
    if (!this.kickerGraphics) {
      this.kickerGraphics = this.add.graphics();
    }
    this.kickerGraphics.clear();

    this.kickers.forEach(kicker => {
      // Green normally, yellow when active
      const color = kicker.active ? 0xFFFF00 : 0x4CAF50;
      this.kickerGraphics.fillStyle(color);
      this.kickerGraphics.fillCircle(kicker.baseX, kicker.baseY, kicker.radius);

      // Arrow indicator
      this.kickerGraphics.fillStyle(0xffffff);
      this.kickerGraphics.fillTriangle(
        kicker.baseX, kicker.baseY - 8,
        kicker.baseX - 6, kicker.baseY + 4,
        kicker.baseX + 6, kicker.baseY + 4
      );
    });
  }

  drawRamps() {
    if (!this.rampGraphics) {
      this.rampGraphics = this.add.graphics();
    }
    this.rampGraphics.clear();

    this.ramps.forEach(ramp => {
      // Cyan color for speed boost ramps
      this.rampGraphics.fillStyle(0x00BCD4);
      this.rampGraphics.fillCircle(ramp.baseX, ramp.baseY, ramp.radius);

      // Speed indicator arrows
      this.rampGraphics.lineStyle(2, 0xffffff);
      this.rampGraphics.strokeCircle(ramp.baseX, ramp.baseY, ramp.radius - 5);
    });
  }

  updateComboTimer() {
    if (this.comboTimer > 0) {
      this.comboTimer--;
      if (this.comboTimer === 0) {
        this.multiplier = 1;
        this.multiplierText.setText('x1');
      }
    }
  }

  checkBallLost() {
    if (!this.ball) return;

    if (this.ball.position.y > 1050) {
      this.ballsLeft--;
      this.ballsText.setText('Balls: ' + this.ballsLeft);

      if (this.ballsLeft > 0) {
        this.resetBall();
      } else {
        this.gameOver();
      }
    }
  }

  resetBall() {
    // Remove old ball
    this.matter.world.remove(this.ball);
    this.ballTrail = [];

    // Create new ball
    this.ball = this.matter.add.circle(730, 850, 10, {
      restitution: 0.6,
      friction: 0.001,
      frictionAir: 0.01,
      label: 'ball'
    });

    this.matter.body.setStatic(this.ball, true);
    this.launched = false;
    this.launchText.setVisible(true);
  }

  gameOver() {
    const gameOverText = this.add.text(400, 500, 'GAME OVER\n\nFinal Score: ' + this.score + '\n\nPress SPACE to restart', {
      fontSize: '32px',
      color: '#ffffff',
      align: 'center'
    });
    gameOverText.setOrigin(0.5);

    // Wait for space to restart
    this.input.keyboard.once('keydown-SPACE', () => {
      this.scene.restart();
    });

    // Remove ball
    if (this.ball) {
      this.matter.world.remove(this.ball);
      this.ball = null;
    }
  }
}
