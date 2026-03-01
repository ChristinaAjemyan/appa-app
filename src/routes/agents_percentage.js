const express = require('express');
const { AgentPercentageService } = require('../services');
const router = express.Router();

// Get all agents percentage records
router.get('/', async (req, res) => {
  try {
    const page = Math.max(1, parseInt(req.query.page) || 1);
    const limit = Math.min(100, Math.max(1, parseInt(req.query.limit) || 10));
    const offset = (page - 1) * limit;

    const { total, records } = await AgentPercentageService.findAllWithCount({ limit, offset });
    res.json({
      data: records,
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) }
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get agent percentage by ID
router.get('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const record = await AgentPercentageService.findById(id);
    res.json(record);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Create a new agent percentage record
router.post('/', async (req, res) => {
  try {
    const { company, agent_code_in, agent_code_not, region_in, region_not, bm_min, bm_max, bm_exact, brand_in, hp_min, hp_max, period, percentage } = req.body;
    
    if (!company || percentage === undefined) {
      res.status(400).json({ error: 'company and percentage are required' });
      return;
    }

    const record = await AgentPercentageService.create({
      company, agent_code_in, agent_code_not, region_in, region_not, bm_min, bm_max, bm_exact,
      brand_in, hp_min, hp_max, period, percentage
    });
    res.status(201).json(record);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Update agent percentage record
router.put('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { company, agent_code_in, agent_code_not, region_in, region_not, bm_min, bm_max, bm_exact, brand_in, hp_min, hp_max, period, percentage } = req.body;

    const record = await AgentPercentageService.update(id, {
      company, agent_code_in, agent_code_not, region_in, region_not, bm_min, bm_max, bm_exact,
      brand_in, hp_min, hp_max, period, percentage
    });
    res.json(record);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Delete agent percentage record
router.delete('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const result = await AgentPercentageService.delete(id);
    res.json(result);
  } catch (err) {
    res.status(err.message.includes('not found') ? 404 : 500).json({ error: err.message });
  }
});

module.exports = router;
