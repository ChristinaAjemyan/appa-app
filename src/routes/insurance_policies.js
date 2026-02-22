const express = require('express');
const { getDatabase } = require('../database');
const router = express.Router();

// Get all insurance policies with sorting, pagination, and filtering
router.get('/', async (req, res) => {
  try {
    const db = await getDatabase();
    
    // Pagination
    const page = Math.max(1, parseInt(req.query.page) || 1);
    const limit = Math.min(1000, Math.max(1, parseInt(req.query.limit) || 10));
    const offset = (page - 1) * limit;
    
    // Sorting
    const sortBy = req.query.sortBy || 'id';
    const sortOrder = (req.query.sortOrder || 'asc').toUpperCase() === 'DESC' ? 'DESC' : 'ASC';
    const allowedColumns = ['id', 'company', 'agent_inner_code', 'polis_number', 'owner_name', 
                           'start_date', 'end_date', 'region', 'price', 'agent_income', 'income'];
    const validSortBy = allowedColumns.includes(sortBy) ? sortBy : 'id';
    
    // Filtering
    let whereConditions = [];
    let params = [];
    
    if (req.query.company) {
      whereConditions.push("company LIKE ?");
      params.push(`%${req.query.company}%`);
    }
    if (req.query.agent_inner_code) {
      whereConditions.push("agent_inner_code LIKE ?");
      params.push(`%${req.query.agent_inner_code}%`);
    }
    if (req.query.polis_number) {
      whereConditions.push("polis_number LIKE ?");
      params.push(`%${req.query.polis_number}%`);
    }
    if (req.query.region) {
      whereConditions.push("region LIKE ?");
      params.push(`%${req.query.region}%`);
    }
    if (req.query.owner_name) {
      whereConditions.push("owner_name LIKE ?");
      params.push(`%${req.query.owner_name}%`);
    }
    if (req.query.minPrice !== undefined) {
      whereConditions.push("price >= ?");
      params.push(parseFloat(req.query.minPrice));
    }
    if (req.query.maxPrice !== undefined) {
      whereConditions.push("price <= ?");
      params.push(parseFloat(req.query.maxPrice));
    }
    
    const whereClause = whereConditions.length > 0 ? `WHERE ${whereConditions.join(' AND ')}` : '';
    
    // Get total count
    const countQuery = `SELECT COUNT(*) as total FROM insurance_policies ${whereClause}`;
    const countResult = await db.get(countQuery, params);
    const total = countResult.total;
    
    // Get paginated and sorted data
    const query = `SELECT * FROM insurance_policies ${whereClause} ORDER BY ${validSortBy} ${sortOrder} LIMIT ? OFFSET ?`;
    const rows = await db.all(query, [...params, limit, offset]);
    
    res.json({
      data: rows,
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
