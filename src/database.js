const sqlite3 = require('sqlite3');
const sqlite = require('sqlite');
const path = require('path');

const dbPath = path.join(__dirname, '../database.db');

let db = null;

// Enable verbose logging for sqlite3
// Uncomment the line below to see all SQL statements
// sqlite3.verbose();

async function getDatabase() {
  if (!db) {
    db = await sqlite.open({
      filename: dbPath,
      driver: sqlite3.Database,
    });

    console.log('Connected to SQLite database');
    
    // Enable query logging
    enableQueryLogging(db);
    
    await initializeDatabase();
  }
  return db;
}

/**
 * Wraps db.get() and db.run() to log SQL queries
 */
function enableQueryLogging(dbInstance) {
  const originalGet = dbInstance.get.bind(dbInstance);
  const originalRun = dbInstance.run.bind(dbInstance);
  const originalAll = dbInstance.all.bind(dbInstance);
  
  dbInstance.get = function(sql, params, ...args) {
    console.log('[SQL GET]', sql, params);
    return originalGet(sql, params, ...args);
  };
  
  dbInstance.run = function(sql, params, ...args) {
    console.log('[SQL RUN]', sql, params);
    return originalRun(sql, params, ...args);
  };
  
  dbInstance.all = function(sql, params, ...args) {
    console.log('[SQL ALL]', sql, params);
    return originalAll(sql, params, ...args);
  };
}

async function initializeDatabase() {
  const database = await getDatabase();

  // Create agents_percentage table
  await database.exec(`
    CREATE TABLE IF NOT EXISTS agents_percentage (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      company TEXT NOT NULL,
      agent_code_in TEXT,
      agent_code_not TEXT,
      region_in TEXT,
      region_not TEXT,
      bm_min INTEGER,
      bm_max INTEGER,
      bm_exact INTEGER,
      brand_in TEXT,
      hp_min INTEGER,
      hp_max INTEGER,
      period TEXT,
      percentage REAL NOT NULL
    )
  `);

  // Create companies_percentage table
  await database.exec(`
    CREATE TABLE IF NOT EXISTS companies_percentage (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      company TEXT NOT NULL,
      agent_code_in TEXT,
      agent_code_not TEXT,
      region_in TEXT,
      region_not TEXT,
      bm_min INTEGER,
      bm_max INTEGER,
      bm_exact INTEGER,
      brand_in TEXT,
      hp_min INTEGER,
      hp_max INTEGER,
      period TEXT,
      percentage REAL NOT NULL
    )
  `);

  // Create agents table
  await database.exec(`
    CREATE TABLE IF NOT EXISTS agents (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      code TEXT,
      name TEXT,
      phone_number TEXT,
      NAIRI TEXT,
      REGO TEXT,
      ARMENIA TEXT,
      SIL TEXT,
      INGO TEXT,
      LIGA TEXT
    )
  `);

  // Create insurance_policies table
  await database.exec(`
    CREATE TABLE IF NOT EXISTS insurance_policies (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      company TEXT,
      agent_company_code TEXT,
      agent_inner_code TEXT,
      agent_name TEXT,
      polis_number TEXT,
      owner_name TEXT,
      start_date TEXT,
      end_date TEXT,
      region TEXT,
      phone_number TEXT,
      bm_class TEXT,
      car_model TEXT,
      car_number TEXT,
      hp TEXT,
      period TEXT,
      info TEXT,
      price REAL,
      agent_percent REAL,
      company_percent REAL,
      agent_income REAL,
      income REAL
    )
  `);

  // Create company table
  await database.exec(`
    CREATE TABLE IF NOT EXISTS company (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT,
      company_percent REAL,
      agent_percent REAL
    )
  `);
}

// Initialize database on module load
getDatabase().catch(err => {
  console.error('Error initializing database:', err);
});

module.exports = { getDatabase };

