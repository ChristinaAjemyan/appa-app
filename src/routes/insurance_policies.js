const express = require('express');
const { InsurancePolicyService } = require('../services');
const router = express.Router();

// Get all insurance policies with sorting, pagination, and filtering
router.get('/', async (req, res) => {
  try {
    const page = Math.max(1, parseInt(req.query.page) || 1);
    const limit = Math.min(1000, Math.max(1, parseInt(req.query.limit) || 10));
    const offset = (page - 1) * limit;
    
    const sortBy = req.query.sortBy || 'id';
    const sortOrder = req.query.sortOrder || 'ASC';
    
    // Build filters object from query params
    const filters = {};
    if (req.query.company) filters.company = req.query.company;
    if (req.query.agent_inner_code) filters.agent_inner_code = req.query.agent_inner_code;
    if (req.query.polis_number) filters.polis_number = req.query.polis_number;
    if (req.query.region) filters.region = req.query.region;
    if (req.query.owner_name) filters.owner_name = req.query.owner_name;
    if (req.query.minPrice !== undefined) filters.minPrice = req.query.minPrice;
    if (req.query.maxPrice !== undefined) filters.maxPrice = req.query.maxPrice;
    if (req.query.startDate) filters.startDate = req.query.startDate;
    if (req.query.endDate) filters.endDate = req.query.endDate;
    
    const { total, policies } = await InsurancePolicyService.search(filters, 
      { limit, offset, sortBy, sortOrder });

    res.json({
      data: policies,
      pagination: {
        page,
        limit,
        total,
        totalPages: Math.ceil(total / limit)
      }
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get insurance policy by ID
router.get('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const policy = await InsurancePolicyService.findById(id);
    res.json(policy);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Create a new insurance policy
router.post('/', async (req, res) => {
  try {
    const { company, agent_company_code, agent_inner_code, agent_name, polis_number, 
            owner_name, start_date, end_date, region, phone_number, bm_class, car_model, 
            car_number, hp, period, info, price, agent_percent, company_percent, 
            agent_income, income } = req.body;
    
    if (!company || !polis_number) {
      res.status(400).json({ error: 'company and polis_number are required' });
      return;
    }

    const policy = await InsurancePolicyService.create({
      company, agent_company_code, agent_inner_code, agent_name, polis_number, owner_name,
      start_date, end_date, region, phone_number, bm_class, car_model, car_number, hp,
      period, info, price, agent_percent, company_percent, agent_income, income
    });
    res.status(201).json(policy); 
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Update insurance policy
router.put('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { company, agent_company_code, agent_inner_code, agent_name, polis_number, 
            owner_name, start_date, end_date, region, phone_number, bm_class, car_model, 
            car_number, hp, period, info, price, agent_percent, company_percent, 
            agent_income, income } = req.body;

    const policy = await InsurancePolicyService.update(id, {
      company, agent_company_code, agent_inner_code, agent_name, polis_number, owner_name,
      start_date, end_date, region, phone_number, bm_class, car_model, car_number, hp,
      period, info, price, agent_percent, company_percent, agent_income, income
    });
    res.json(policy);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Delete insurance policy
router.delete('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const result = await InsurancePolicyService.delete(id);
    res.json(result);
  } catch (err) {
    res.status(err.message.includes('not found') ? 404 : 500).json({ error: err.message });
  }
});

module.exports = router;
