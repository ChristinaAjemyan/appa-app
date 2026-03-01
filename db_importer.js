const fs = require('fs');
const path = require('path');
const csv = require('csv-parser');
const { AgentService, AgentPercentageService, CompanyService, CompanyPercentageService } = require('./src/services');

const csvDir = path.join(__dirname, 'csv');

// Map of CSV filenames to their corresponding services
const csvMapping = {
  'agents.csv': {
    service: AgentService,
    name: 'agents'
  },
  'agents_percentage.csv': {
    service: AgentPercentageService,
    name: 'agents_percentage'
  },
  'companies_percentage.csv': {
    service: CompanyPercentageService,
    name: 'companies_percentage'
  },
  'companies.csv': {
    service: CompanyService,
    name: 'companies'
  },
};

/**
 * Import CSV file into database
 * Inserts all records using the appropriate service
 */
async function importCsvFile(filename, config) {
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
          const stats = await processRecords(records, config);
          console.log(`✅ ${filename} import complete:`);
          console.log(`   - Processed: ${stats.processed}`);
          console.log(`   - Inserted: ${stats.inserted}`);
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
 * Process records and insert them using the service
 */
async function processRecords(records, config) {
  let processed = 0;
  let inserted = 0;
  let errors = 0;

  if (records.length === 0) {
    return { processed: 0, inserted: 0, errors: 0 };
  }

  for (const record of records) {
    // Convert NULL strings to actual null
    const cleanRecord = {};
    for (const [key, value] of Object.entries(record)) {
      cleanRecord[key] = value === 'NULL' ? null : value;
    }

    try {
      await config.service.upsert(cleanRecord, cleanRecord);
      inserted++;
    } catch (err) {
      console.error(`Error processing record in ${config.name}:`, err);
      errors++;
    }
    
    processed++;
  }

  return { processed, inserted, errors };
}

/**
 * Run all imports
 */
async function runImports() {
  console.log('🚀 Starting CSV import process...\n');

  try {
    for (const [filename, config] of Object.entries(csvMapping)) {
      await importCsvFile(filename, config);
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
