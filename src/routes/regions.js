const express = require('express');
const { RegionService } = require('../services');
const router = express.Router();

// Get all regions
router.get('/', async (req, res) => {
  try {
    const regions = await RegionService.findAll();
    res.json({ data: regions });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get region by ID
router.get('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const region = await RegionService.findById(id);
    res.json(region);
  } catch (err) {
    res.status(err.message.includes('not found') ? 404 : 500).json({ error: err.message });
  }
});

// Create a new region
router.post('/', async (req, res) => {
  try {
    const { region_code, name } = req.body;

    if (!region_code || !name) {
      res.status(400).json({ error: 'region_code and name are required' });
      return;
    }

    const region = await RegionService.create({
      region_code,
      name
    });
    res.status(201).json(region);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Update region
router.put('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { region_code, name } = req.body;

    const region = await RegionService.update(id, {
      region_code,
      name
    });
    res.json(region);
  } catch (err) {
    res.status(err.message.includes('not found') ? 404 : 500).json({ error: err.message });
  }
});

// Delete region
router.delete('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const result = await RegionService.delete(id);
    res.json(result);
  } catch (err) {
    res.status(err.message.includes('not found') ? 404 : 500).json({ error: err.message });
  }
});

module.exports = router;
