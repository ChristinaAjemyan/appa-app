const { Sequelize } = require('sequelize');

const sequelize = new Sequelize({
  dialect: 'postgres',
  host: process.env.DB_HOST || 'localhost',
  port: process.env.DB_PORT || 5432,
  username: process.env.DB_USER || 'postgres',
  password: process.env.DB_PASSWORD || 'password',
  database: process.env.DB_NAME || 'appa',
  logging: process.env.DB_LOGGING === 'true' ? console.log : false,
  define: {
    timestamps: false
  }
});

const db = {};

// Import models
db.Agent = require('./models/Agent')(sequelize);
db.AgentPercentage = require('./models/AgentPercentage')(sequelize);
db.Company = require('./models/Company')(sequelize);
db.CompanyPercentage = require('./models/CompanyPercentage')(sequelize);
db.InsurancePolicy = require('./models/InsurancePolicy')(sequelize);
db.User = require('./models/User')(sequelize);

db.sequelize = sequelize;
db.Sequelize = Sequelize;

/**
 * Initialize database connection and sync models
 */
async function getDatabase() {
  try {
    await sequelize.authenticate();
    console.log('✓ PostgreSQL connection successful');
    
    // Sync models with database
    await sequelize.sync({ alter: true });
    console.log('✓ Database models synchronized');
    
    return db;
  } catch (err) {
    console.error('✗ Error connecting to database:', err.message);
    throw err;
  }
}

// Initialize database on module load
getDatabase().catch(err => {
  console.error('Error initializing database:', err);
});

// Graceful shutdown
process.on('exit', async () => {
  await sequelize.close();
  console.log('Database connection closed');
});

module.exports = { getDatabase, db, sequelize };

