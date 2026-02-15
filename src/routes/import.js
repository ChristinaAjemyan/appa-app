const express = require('express');
const { spawn } = require('child_process');
const path = require('path');
const router = express.Router();

/**
 * POST /api/import
 * Trigger CSV import process
 */
router.post('/', (req, res) => {
  const importer = spawn('node', [path.join(__dirname, '../db_importer.js')]);
  let output = '';
  let error = '';

  importer.stdout.on('data', (data) => {
    output += data.toString();
  });

  importer.stderr.on('data', (data) => {
    error += data.toString();
  });

  importer.on('close', (code) => {
    if (code === 0) {
      res.json({
        success: true,
        message: 'Import completed successfully',
        output: output
      });
    } else {
      res.status(500).json({
        success: false,
        message: 'Import failed',
        error: error
      });
    }
  });
});

module.exports = router;
