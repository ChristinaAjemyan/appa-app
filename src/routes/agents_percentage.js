const express = require('express');
const { getDatabase } = require('../database');
const router = express.Router();

// Get all agents percentage records
router.get('/', async (req, res) => {
  try {
    const db = await getDatabase();
    const rows = await db.all('SELECT * FROM agents_percentage');
    res.json(rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get agent percentage by ID
router.get('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const db = await getDatabase();
    const row = await db.get('SELECT * FROM agents_percentage WHERE id = ?', [id]);
    
    if (!row) {
      res.status(404).json({ error: 'Record not found' });
      return;
    }
    res.json(row);
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

    const db = await getDatabase();
    const result = await db.run(
      'INSERT INTO agents_percentage (company, agent_code_in, agent_code_not, region_in, region_not, bm_min, bm_max, bm_exact, brand_in, hp_min, hp_max, period, percentage) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
      [company, agent_code_in, agent_code_not, region_in, region_not, bm_min, bm_max, bm_exact, brand_in, hp_min, hp_max, period, percentage]
    );
    
    res.status(201).json({ id: result.lastID, company, agent_code_in, agent_code_not, region_in, region_not, bm_min, bm_max, bm_exact, brand_in, hp_min, hp_max, period, percentage });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Update agent percentage record
router.put('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { company, agent_code_in, agent_code_not, region_in, region_not, bm_min, bm_max, bm_exact, brand_in, hp_min, hp_max, period, percentage } = req.body;

    const db = await getDatabase();
    const result = await db.run(
      'UPDATE agents_percentage SET company = ?, agent_code_in = ?, agent_code_not = ?, region_in = ?, region_not = ?, bm_min = ?, bm_max = ?, bm_exact = ?, brand_in = ?, hp_min = ?, hp_max = ?, period = ?, percentage = ? WHERE id = ?',
      [company, agent_code_in, agent_code_not, region_in, region_not, bm_min, bm_max, bm_exact, brand_in, hp_min, hp_max, period, percentage, id]
    );

    if (result.changes === 0) {
      res.status(404).json({ error: 'Record not found' });
      return;
    }
    
    res.json({ id, company, agent_code_in, agent_code_not, region_in, region_not, bm_min, bm_max, bm_exact, brand_in, hp_min, hp_max, period, percentage });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Delete agent percentage record
router.delete('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const db = await getDatabase();
    const result = await db.run('DELETE FROM agents_percentage WHERE id = ?', [id]);

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
