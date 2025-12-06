// ABOUTME: Boot scene for loading assets and initializing game
// ABOUTME: First scene that runs, handles preloading

import Phaser from 'phaser';

export default class BootScene extends Phaser.Scene {
  constructor() {
    super('BootScene');
  }

  preload() {
    // Display loading text
    const loadingText = this.add.text(400, 500, 'Loading...', {
      fontSize: '32px',
      color: '#ffffff'
    });
    loadingText.setOrigin(0.5);

    // Load any assets here in the future
    // this.load.image('ball', 'assets/images/ball.png');
  }

  create() {
    // Start the editor scene
    this.scene.start('EditorScene');
  }
}
