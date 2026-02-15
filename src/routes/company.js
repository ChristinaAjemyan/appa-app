const express = require('express');
const { getDatabase } = require('../database');
const router = express.Router();

// Get all companies
router.get('/', async (req, res) => {
  try {
    const db = await getDatabase();
    const rows = await db.all('SELECT * FROM company');
    res.json(rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get company by ID
router.get('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const db = await getDatabase();
    const row = await db.get('SELECT * FROM company WHERE id = ?', [id]);
    
    if (!row) {
      res.status(404).json({ error: 'Company not found' });
      return;
    }
    res.json(row);
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

    const db = await getDatabase();
    const result = await db.run(
      'INSERT INTO company (name, company_percent, agent_percent) VALUES (?, ?, ?)',
      [name, company_percent, agent_percent]
    );
    
    res.status(201).json({ id: result.lastID, name, company_percent, agent_percent });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Update company
router.put('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { name, company_percent, agent_percent } = req.body;

    const db = await getDatabase();
    const result = await db.run(
      'UPDATE company SET name = ?, company_percent = ?, agent_percent = ? WHERE id = ?',
      [name, company_percent, agent_percent, id]
    );

    if (result.changes === 0) {
      res.status(404).json({ error: 'Company not found' });
      return;
    }
    
    res.json({ id, name, company_percent, agent_percent });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Delete company
router.delete('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const db = await getDatabase();
    const result = await db.run('DELETE FROM company WHERE id = ?', [id]);

    if (result.changes === 0) {
      res.status(404).json({ error: 'Company not found' });
      return;
    }
    
    res.json({ message: 'Company deleted successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;
