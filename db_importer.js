const fs = require('fs');
const path = require('path');
const csv = require('csv-parser');
const { getDatabase } = require('./src/database');

const csvDir = path.join(__dirname, 'csv');

// Map of CSV filenames to their corresponding database tables
const csvMapping = {
  'agents.csv': {
    table: 'agents',
    idColumn: 'id'
  },
  'agents_percentage.csv': {
    table: 'agents_percentage',
    idColumn: 'id'
  },
  'companies_percentage.csv': {
    table: 'companies_percentage',
    idColumn: 'id'
  },
  'companies.csv': {
    table: 'company',
    idColumn: 'id'
  },
};

/**
 * Import CSV file into database
 * Updates existing records with same id, inserts new ones
 */
async function importCsvFile(filename, tableConfig) {
  const filePath = path.join(csvDir, filename);
  
  if (!fs.existsSync(filePath)) {
    console.warn(`⚠️  File not found: ${filePath}`);
    return;
  }

  return new Promise((resolve, reject) => {
    const records = [];

    fs.createReadStream(filePath)
      .pipe(csv())
      .on('data', (row) => {
        records.push(row);
      })
      .on('end', async () => {
        try {
          const stats = await processRecords(records, tableConfig);
          console.log(`✅ ${filename} import complete:`);
          console.log(`   - Processed: ${stats.processed}`);
          console.log(`   - Inserted: ${stats.inserted}`);
          console.log(`   - Updated: ${stats.updated}`);
          console.log(`   - Errors: ${stats.errors}`);
          resolve();
        } catch (err) {
          reject(err);
        }
      })
      .on('error', (err) => {
        console.error(`❌ Error reading ${filename}:`, err.message);
        reject(err);
      });
  });
}

/**
 * Process records and insert/update them
 */
async function processRecords(records, tableConfig) {
  let processed = 0;
  let inserted = 0;
  let updated = 0;
  let errors = 0;

  if (records.length === 0) {
    return { processed: 0, inserted: 0, updated: 0, errors: 0 };
  }

  const db = await getDatabase();

  for (const record of records) {
    // Convert NULL strings to actual null
    const cleanRecord = {};
    for (const [key, value] of Object.entries(record)) {
      cleanRecord[key] = value === 'NULL' ? null : value;
    }

    const recordId = cleanRecord[tableConfig.idColumn];

    try {
      // Check if record exists
      const existingRecord = await db.get(
        `SELECT id FROM ${tableConfig.table} WHERE ${tableConfig.idColumn} = ?`,
        [recordId]
      );

      if (existingRecord) {
        // Update existing record
        const success = await updateRecord(db, tableConfig.table, cleanRecord, tableConfig.idColumn);
        if (success) updated++;
        else errors++;
      } else {
        // Insert new record
        const success = await insertRecord(db, tableConfig.table, cleanRecord);
        if (success) inserted++;
        else errors++;
      }
    } catch (err) {
      console.error(`Error processing record in ${tableConfig.table}:`, err.message);
      errors++;
    }
    
    processed++;
  }

  return { processed, inserted, updated, errors };
}

/**
 * Insert a new record
 */
async function insertRecord(db, table, record) {
  const columns = Object.keys(record);
  const placeholders = columns.map(() => '?').join(', ');
  const values = columns.map(col => record[col]);

  const sql = `INSERT INTO ${table} (${columns.join(', ')}) VALUES (${placeholders})`;

  try {
    await db.run(sql, values);
    return true;
  } catch (err) {
    console.error(`Error inserting into ${table}:`, err.message);
    return false;
  }
}

/**
 * Update an existing record
 */
async function updateRecord(db, table, record, idColumn) {
  const recordId = record[idColumn];
  const columns = Object.keys(record).filter(col => col !== idColumn);
  const setClause = columns.map(col => `${col} = ?`).join(', ');
  const values = columns.map(col => record[col]);
  values.push(recordId);

  const sql = `UPDATE ${table} SET ${setClause} WHERE ${idColumn} = ?`;

  try {
    await db.run(sql, values);
    return true;
  } catch (err) {
    console.error(`Error updating ${table}:`, err.message);
    return false;
  }
}

/**
 * Run all imports
 */
async function runImports() {
  console.log('🚀 Starting CSV import process...\n');

  try {
    for (const [filename, tableConfig] of Object.entries(csvMapping)) {
      await importCsvFile(filename, tableConfig);
    }
    console.log('\n✨ All imports completed successfully!');
    process.exit(0);
  } catch (err) {
    console.error('\n❌ Import process failed:', err);
    process.exit(1);
  }
}

// Run the import
runImports();
