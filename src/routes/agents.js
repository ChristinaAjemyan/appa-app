const express = require('express');
const { AgentService } = require('../services');
const router = express.Router();

// Get all agents
router.get('/', async (req, res) => {
  try {
    const page = Math.max(1, parseInt(req.query.page) || 1);
    const limit = Math.min(100, Math.max(1, parseInt(req.query.limit) || 10));
    const offset = (page - 1) * limit;
    console.log(`Fetching agents - Page: ${page}, Limit: ${limit}, Offset: ${offset}`);
    const { total, agents } = await AgentService.findAllWithCount({ limit, offset });
    res.json({
      data: agents,
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) }
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get agent by ID
router.get('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const agent = await AgentService.findById(id);
    res.json(agent);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Create a new agent
router.post('/', async (req, res) => {
  try {
    const { code, name, phone_number, NAIRI, REGO, ARMENIA, SIL, INGO, LIGA } = req.body;

    if (!code || !name) {
      res.status(400).json({ error: 'Code and name are required' });
      return;
    }

    const agent = await AgentService.create({
      code, name, phone_number, NAIRI, REGO, ARMENIA, SIL, INGO, LIGA
    });
    res.status(201).json(agent);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Update agent
router.put('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { code, name, phone_number, NAIRI, REGO, ARMENIA, SIL, INGO, LIGA } = req.body;

    const agent = await AgentService.update(id, {
      code, name, phone_number, NAIRI, REGO, ARMENIA, SIL, INGO, LIGA
    });
    res.json(agent);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Delete agent
router.delete('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const result = await AgentService.delete(id);
    res.json(result);
  } catch (err) {
    res.status(err.message.includes('not found') ? 404 : 500).json({ error: err.message });
  }
});

module.exports = router;
