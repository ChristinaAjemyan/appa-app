const express = require('express');
const multer = require('multer');
const csv = require('csv-parser');
const { Readable } = require('stream');
const { getDatabase } = require('../database');
const router = express.Router();

// Configure multer for file uploads
const upload = multer({ storage: multer.memoryStorage() });

/**
 * POST /api/calculate
 * Upload and parse CSV file, insert data into insurance_policies table
 */
router.post('/', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      res.status(400).json({ error: 'No file uploaded' });
      return;
    }

    if (req.file.mimetype !== 'text/csv' && !req.file.originalname.endsWith('.csv')) {
      res.status(400).json({ error: 'File must be a CSV file' });
      return;
    }

    const records = [];

    // Parse CSV from buffer
    await new Promise((resolve, reject) => {
      Readable.from([req.file.buffer])
        .pipe(csv())
        .on('data', (row) => {
          records.push(row);
        })
        .on('end', resolve)
        .on('error', reject);
    });

    if (records.length === 0) {
      res.status(400).json({ error: 'CSV file is empty' });
      return;
    }

    // Insert records into database with calculated percentages
    const stats = await insertRecords(records);
    
    res.json({
      success: true,
      message: 'CSV data imported successfully',
      stats: []
    });
  } catch (err) {
    console.error('Error processing CSV:', err.message);
    res.status(500).json({ error: 'Error processing CSV file: ' + err.message });
  }
});


/**
 * Calculate agent_percent for a record
 * Returns the percentage from agents_percentage table if exceptions match,
 * otherwise returns the default agent_percent from company table
 */
async function calculateAgentPercent(db, record) {
  const { company, agent_inner_code, region, bm_class, car_model, hp, period } = record;
  console.log('Calculating agent_percent for record:', record);
  if (!company) {
    return null;
  }

  // Get default agent_percent from company table
  const companyData = await db.get(
    'SELECT agent_percent FROM company WHERE lower(name) = lower(?)',
    [company]
  );
  console.log('Company data:', companyData);  
  const defaultPercent = companyData?.agent_percent || null;

  // Convert bm_class and hp to numbers for range comparison
  const bmClassNum = bm_class ? parseInt(bm_class, 10) : null;
  const hpNum = hp ? parseInt(hp, 10) : null;
  console.log('BM Class number:', bmClassNum);
  console.log('HP number:', hpNum);
  
  // Check for exceptions in agents_percentage table
  // Match based on: company, agent codes (in/not), regions (in/not), bm_class range, 
  // hp range, brand_in, period
  // Note: region_in and region_not may contain comma-separated values
  const exception = await db.get(
    `SELECT percentage FROM agents_percentage 
     WHERE lower(company) = lower(?) 
     AND (agent_code_in IS NULL OR lower(agent_code_in) LIKE lower(?))
     AND (agent_code_not IS NULL OR lower(agent_code_not) NOT LIKE lower(?))
     AND (region_in IS NULL OR lower(region_in) LIKE lower(?))
     AND (region_not IS NULL OR lower(region_not) NOT LIKE lower(?))
     AND (bm_exact IS NULL OR bm_exact = ?)
     AND ((bm_min IS NULL OR bm_max IS NULL) OR (? >= bm_min AND ? <= bm_max))
     AND ((hp_min IS NULL OR hp_max IS NULL) OR (? >= hp_min AND ? <= hp_max))
     AND (brand_in IS NULL OR lower(brand_in) LIKE lower(?))
     AND (period IS NULL OR lower(period) = lower(?))
     LIMIT 1`,
    [
      company,
      agent_inner_code ? `%${agent_inner_code}%` : null,
      agent_inner_code ? `%${agent_inner_code}%` : null,
      region ? `%${region}%` : null,
      region ? `%${region}%` : null,
      bmClassNum,
      bmClassNum,
      bmClassNum,
      hpNum,
      hpNum,
      car_model ? `%${car_model}%` : null,
      period || null
    ]
  );
  console.log('Exception found:', exception)
  const percentage = exception?.percentage || defaultPercent;
  // Return just the percentage value, not the calculated income
  return percentage;
}

/**
 * Calculate company_percent for a record
 * Returns the percentage from companies_percentage table if exceptions match,
 * otherwise returns the default company_percent from company table
 */
async function calculateCompanyPercent(db, record) {
  const { company, agent_inner_code, region, bm_class, car_model, hp, period } = record;
  console.log('Calculating company_percent for record:', record);
  if (!company) {
    return null;
  }

  // Get default company_percent from company table
  const companyData = await db.get(
    'SELECT company_percent FROM company WHERE lower(name) = lower(?)',
    [company]
  );
  console.log('Company data:', companyData);  
  const defaultPercent = companyData?.company_percent || null;

  // Convert bm_class and hp to numbers for range comparison
  const bmClassNum = bm_class ? parseInt(bm_class, 10) : null;
  const hpNum = hp ? parseInt(hp, 10) : null;
  console.log('BM Class number:', bmClassNum);
  console.log('HP number:', hpNum);
  
  // Check for exceptions in companies_percentage table
  // Match based on: company, agent codes (in/not), regions (in/not), bm_class range, 
  // hp range, brand_in, period
  // Note: region_in and region_not may contain comma-separated values
  const exception = await db.get(
    `SELECT percentage FROM companies_percentage 
     WHERE lower(company) = lower(?) 
     AND (agent_code_in IS NULL OR lower(agent_code_in) LIKE lower(?))
     AND (agent_code_not IS NULL OR lower(agent_code_not) NOT LIKE lower(?))
     AND (region_in IS NULL OR lower(region_in) LIKE lower(?))
     AND (region_not IS NULL OR lower(region_not) NOT LIKE lower(?))
     AND (bm_exact IS NULL OR bm_exact = ?)
     AND ((bm_min IS NULL OR bm_max IS NULL) OR (? >= bm_min AND ? <= bm_max))
     AND ((hp_min IS NULL OR hp_max IS NULL) OR (? >= hp_min AND ? <= hp_max))
     AND (brand_in IS NULL OR lower(brand_in) LIKE lower(?))
     AND (period IS NULL OR lower(period) = lower(?))
     LIMIT 1`,
    [
      company,
      agent_inner_code ? `%${agent_inner_code}%` : null,
      agent_inner_code ? `%${agent_inner_code}%` : null,
      region ? `%${region}%` : null,
      region ? `%${region}%` : null,
      bmClassNum,
      bmClassNum,
      bmClassNum,
      hpNum,
      hpNum,
      car_model ? `%${car_model}%` : null,
      period || null
    ]
  );
  console.log('Exception found:', exception)
  const percentage = exception?.percentage || defaultPercent;
  // Return just the percentage value, not the calculated income
  return percentage;
}

async function insertRecords(records) {
  const db = await getDatabase();
  let inserted = 0;
  let errors = 0;
  console.log('Inserting records into database:', records.length);
  for (const record of records) {
    try {
      // Convert empty strings to null
      const cleanRecord = {};
      for (const [key, value] of Object.entries(record)) {
        cleanRecord[key] = value === '' || value === 'NULL' ? null : value;
      }

      const {
        company,
        agent_company_code,
        agent_inner_code,
        agent_name,
        polis_number,
        owner_name,
        start_date,
        end_date,
        region,
        phone_number,
        bm_class,
        car_model,
        car_number,
        hp,
        period,
        info,
        price
      } = cleanRecord;

      // Calculate agent_percent and company_percent
      const agent_percent = await calculateAgentPercent(db, cleanRecord);
      const company_percent = await calculateCompanyPercent(db, cleanRecord);
      console.log('Calculated agent_percent:', agent_percent);
      console.log('Calculated company_percent:', company_percent);
      const agent_income = price && agent_percent ? (price * agent_percent) / 100 : null;
      const company_income = price && company_percent ? (price * company_percent) / 100 : null;
      const income = agent_income && company_income ? company_income - agent_income : null;

      await db.run(
        `INSERT INTO insurance_policies (
          company, agent_company_code, agent_inner_code, agent_name,
          polis_number, owner_name, start_date, end_date, region,
          phone_number, bm_class, car_model, car_number, hp, period, info, price, agent_percent, company_percent, agent_income, income
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [
          company, agent_company_code, agent_inner_code, agent_name,
          polis_number, owner_name, start_date, end_date, region,
          phone_number, bm_class, car_model, car_number, hp, period, info, price, agent_percent, company_percent, agent_income, income
        ]
      );
      
      inserted++;
    } catch (err) {
      console.error('Error inserting record:', err.message);
      errors++;
    }
  }

  return {
    processed: records.length,
    inserted,
    errors
  };
}

module.exports = router;
