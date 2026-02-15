const express = require('express');
const cors = require('cors');
require('dotenv').config();

const app = express();

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

// Routes
app.use('/api/agents-percentage', require('./src/routes/agents_percentage'));
app.use('/api/companies-percentage', require('./src/routes/companies_percentage'));
app.use('/api/agents', require('./src/routes/agents'));
app.use('/api/company', require('./src/routes/company'));
app.use('/api/import', require('./src/routes/import'));
app.use('/api/calculate', require('./src/routes/calculate'));
app.use('/api/insurance-policies', require('./src/routes/insurance_policies'));

// Health check route
app.get('/api/health', (req, res) => {
  res.json({ status: 'Server is running' });
});

// Error handling middleware
app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).json({ error: 'Internal server error' });
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
