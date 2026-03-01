const express = require('express');
const { CompanyPercentageService } = require('../services');
const router = express.Router();

// Get all companies percentage records
router.get('/', async (req, res) => {
  try {
    const page = Math.max(1, parseInt(req.query.page) || 1);
    const limit = Math.min(100, Math.max(1, parseInt(req.query.limit) || 10));
    const offset = (page - 1) * limit;

    const { total, records } = await CompanyPercentageService.findAllWithCount({ limit, offset });
    res.json({
      data: records,
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) }
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get companies percentage by ID
router.get('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const record = await CompanyPercentageService.findById(id);
    res.json(record);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Create a new companies percentage record
router.post('/', async (req, res) => {
  try {
    const { company, agent_code_in, agent_code_not, region_in, region_not, bm_min, bm_max, bm_exact, brand_in, hp_min, hp_max, period, percentage } = req.body;
    
    if (!company || percentage === undefined) {
      res.status(400).json({ error: 'company and percentage are required' });
      return;
    }

    const record = await CompanyPercentageService.create({
      company, agent_code_in, agent_code_not, region_in, region_not, bm_min, bm_max, bm_exact,
      brand_in, hp_min, hp_max, period, percentage
    });
    res.status(201).json(record);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Update companies percentage record
router.put('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { company, agent_code_in, agent_code_not, region_in, region_not, bm_min, bm_max, bm_exact, brand_in, hp_min, hp_max, period, percentage } = req.body;

    const record = await CompanyPercentageService.update(id, {
      company, agent_code_in, agent_code_not, region_in, region_not, bm_min, bm_max, bm_exact,
      brand_in, hp_min, hp_max, period, percentage
    });
    res.json(record);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Delete companies percentage record
router.delete('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const result = await CompanyPercentageService.delete(id);
    res.json(result);
  } catch (err) {
    res.status(err.message.includes('not found') ? 404 : 500).json({ error: err.message });
  }
});

module.exports = router;
