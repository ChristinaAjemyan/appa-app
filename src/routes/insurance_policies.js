const express = require('express');
const { getDatabase } = require('../database');
const router = express.Router();

// Get all insurance policies
router.get('/', async (req, res) => {
  try {
    const db = await getDatabase();
    const rows = await db.all('SELECT * FROM insurance_policies');
    res.json(rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get insurance policy by ID
router.get('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const db = await getDatabase();
    const row = await db.get('SELECT * FROM insurance_policies WHERE id = ?', [id]);
    
    if (!row) {
      res.status(404).json({ error: 'Record not found' });
      return;
    }
    res.json(row);
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

    const db = await getDatabase();
    const result = await db.run(
      `INSERT INTO insurance_policies (company, agent_company_code, agent_inner_code, 
       agent_name, polis_number, owner_name, start_date, end_date, region, phone_number, 
       bm_class, car_model, car_number, hp, period, info, price, agent_percent, 
       company_percent, agent_income, income) 
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [company, agent_company_code, agent_inner_code, agent_name, polis_number, owner_name, 
       start_date, end_date, region, phone_number, bm_class, car_model, car_number, hp, 
       period, info, price, agent_percent, company_percent, agent_income, income]
    );
    
    res.status(201).json({ 
      id: result.lastID, 
      company, agent_company_code, agent_inner_code, agent_name, polis_number, owner_name, 
      start_date, end_date, region, phone_number, bm_class, car_model, car_number, hp, 
      period, info, price, agent_percent, company_percent, agent_income, income 
    });
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

    const db = await getDatabase();
    const result = await db.run(
      `UPDATE insurance_policies SET company = ?, agent_company_code = ?, 
       agent_inner_code = ?, agent_name = ?, polis_number = ?, owner_name = ?, 
       start_date = ?, end_date = ?, region = ?, phone_number = ?, bm_class = ?, 
       car_model = ?, car_number = ?, hp = ?, period = ?, info = ?, price = ?, 
       agent_percent = ?, company_percent = ?, agent_income = ?, income = ? 
       WHERE id = ?`,
      [company, agent_company_code, agent_inner_code, agent_name, polis_number, owner_name, 
       start_date, end_date, region, phone_number, bm_class, car_model, car_number, hp, 
       period, info, price, agent_percent, company_percent, agent_income, income, id]
    );

    if (result.changes === 0) {
      res.status(404).json({ error: 'Record not found' });
      return;
    }
    
    res.json({ id, company, agent_company_code, agent_inner_code, agent_name, polis_number, 
               owner_name, start_date, end_date, region, phone_number, bm_class, car_model, 
               car_number, hp, period, info, price, agent_percent, company_percent, 
               agent_income, income });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Delete insurance policy
router.delete('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const db = await getDatabase();
    const result = await db.run('DELETE FROM insurance_policies WHERE id = ?', [id]);

    if (result.changes === 0) {
      res.status(404).json({ error: 'Record not found' });
      return;
    }
    
    res.json({ message: 'Record deleted successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;
