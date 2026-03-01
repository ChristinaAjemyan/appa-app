const express = require('express');
const { CompanyService } = require('../services');
const router = express.Router();

// Get all companies
router.get('/', async (req, res) => {
  try {
    const page = Math.max(1, parseInt(req.query.page) || 1);
    const limit = Math.min(100, Math.max(1, parseInt(req.query.limit) || 10));
    const offset = (page - 1) * limit;

    const { total, companies } = await CompanyService.findAllWithCount({ limit, offset });
    res.json({
      data: companies,
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) }
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get company by ID
router.get('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const company = await CompanyService.findById(id);
    res.json(company);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Create a new company
router.post('/', async (req, res) => {
  try {
    const { name, company_percent, agent_percent } = req.body;

    if (!name) {
      res.status(400).json({ error: 'Name is required' });
      return;
    }

    const company = await CompanyService.create({
      name, company_percent, agent_percent
    });
    res.status(201).json(company);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Update company
router.put('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { name, company_percent, agent_percent } = req.body;

    const company = await CompanyService.update(id, {
      name, company_percent, agent_percent
    });
    res.json(company);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Delete company
router.delete('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const result = await CompanyService.delete(id);
    res.json(result);
  } catch (err) {
    res.status(err.message.includes('not found') ? 404 : 500).json({ error: err.message });
  }
});

module.exports = router;
