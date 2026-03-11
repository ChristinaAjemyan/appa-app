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
db.Region = require('./models/Region')(sequelize);

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
    
    // Seed regions data if table is empty
    const regionCount = await db.Region.count();
    if (regionCount === 0) {
      const regionsData = [
        { region_code: 'SH', name: 'Շիրակ' },
        { region_code: 'YR', name: 'Երևան' },
        { region_code: 'KT', name: 'Կոտայք' },
        { region_code: 'AR', name: 'Արարատ' },
        { region_code: 'AM', name: 'Արմավիր' },
        { region_code: 'AG', name: 'Արագածոտն' },
        { region_code: 'LR', name: 'Լոռի' },
        { region_code: 'TV', name: 'Տավուշ' },
        { region_code: 'SY', name: 'Սյունիք' },
        { region_code: 'VD', name: 'Վայոց Ձոր' },
        { region_code: 'GE', name: 'Գեղարքունիկ' }
      ];
      await db.Region.bulkCreate(regionsData);
      console.log('✓ Regions seeded successfully');
    }
    
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

