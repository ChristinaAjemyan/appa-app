const express = require('express');
const { getDatabase } = require('../database');
const router = express.Router();

// Get all agents
router.get('/', async (req, res) => {
  try {
    const db = await getDatabase();
    const rows = await db.all('SELECT * FROM agents');
    res.json(rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get agent by ID
router.get('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const db = await getDatabase();
    const row = await db.get('SELECT * FROM agents WHERE id = ?', [id]);
    
    if (!row) {
      res.status(404).json({ error: 'Agent not found' });
      return;
    }
    res.json(row);
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

    const db = await getDatabase();
    const result = await db.run(
      'INSERT INTO agents (code, name, phone_number, NAIRI, REGO, ARMENIA, SIL, INGO, LIGA) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)',
      [code, name, phone_number, NAIRI, REGO, ARMENIA, SIL, INGO, LIGA]
    );
    
    res.status(201).json({ id: result.lastID, code, name, phone_number, NAIRI, REGO, ARMENIA, SIL, INGO, LIGA });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Update agent
router.put('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { code, name, phone_number, NAIRI, REGO, ARMENIA, SIL, INGO, LIGA } = req.body;

    const db = await getDatabase();
    const result = await db.run(
      'UPDATE agents SET code = ?, name = ?, phone_number = ?, NAIRI = ?, REGO = ?, ARMENIA = ?, SIL = ?, INGO = ?, LIGA = ? WHERE id = ?',
      [code, name, phone_number, NAIRI, REGO, ARMENIA, SIL, INGO, LIGA, id]
    );

    if (result.changes === 0) {
      res.status(404).json({ error: 'Agent not found' });
      return;
    }
    
    res.json({ id, code, name, phone_number, NAIRI, REGO, ARMENIA, SIL, INGO, LIGA });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Delete agent
router.delete('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const db = await getDatabase();
    const result = await db.run('DELETE FROM agents WHERE id = ?', [id]);

    if (result.changes === 0) {
      res.status(404).json({ error: 'Agent not found' });
      return;
    }
    
    res.json({ message: 'Agent deleted successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;
