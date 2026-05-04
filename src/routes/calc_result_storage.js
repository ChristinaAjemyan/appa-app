const express = require('express');
const { Op } = require('sequelize');
const { db } = require('../database');
const router = express.Router();

// GET /api/calc-result-storage?prefix=xxx
router.get('/', async (req, res) => {
  try {
    const { prefix } = req.query;
    const where = prefix ? { key: { [Op.like]: `${prefix}%` } } : {};
    const records = await db.CalcResultStorage.findAll({ where, attributes: ['key'] });
    res.json({ keys: records.map(r => r.key) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// GET /api/calc-result-storage/:key
router.get('/:key', async (req, res) => {
  try {
    const record = await db.CalcResultStorage.findByPk(req.params.key);
    if (!record) {
      return res.status(404).json({ error: 'Key not found' });
    }
    res.json({ key: record.key, value: record.value });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// PUT /api/calc-result-storage/:key
router.put('/:key', async (req, res) => {
  try {
    const { value } = req.body;
    await db.CalcResultStorage.upsert({ key: req.params.key, value });
    res.json({ ok: true });
  } catch (err) {
    console.error('calc_result_storage PUT error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;
