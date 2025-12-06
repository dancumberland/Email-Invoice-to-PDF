// ABOUTME: Phaser game initialization and configuration
// ABOUTME: Entry point for the Dream Pinball game

import Phaser from 'phaser';
import BootScene from './scenes/BootScene.js';
import GameScene from './scenes/GameScene.js';
import EditorScene from './scenes/EditorScene.js';

const config = {
  type: Phaser.WEBGL,
  width: 800,
  height: 1000,
  parent: 'game-container',
  backgroundColor: '#2a2a2a',
  physics: {
    default: 'matter',
    matter: {
      gravity: { y: 0.5 },
      debug: true
    }
  },
  scene: [BootScene, EditorScene, GameScene]
};

const game = new Phaser.Game(config);
