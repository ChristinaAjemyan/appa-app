const { db } = require('../database');

class CompanyService {
  /**
   * Get all companies with optional pagination
   */
  static async findAll(options = {}) {
    try {
      const { limit = 10, offset = 0 } = options;
      const companies = await db.Company.findAll({
        limit,
        offset,
        order: [['id', 'ASC']]
      });
      return companies;
    } catch (error) {
      throw new Error(`Error fetching companies: ${error.message}`);
    }
  }

  /**
   * Get all companies with count for pagination
   */
  static async findAllWithCount(options = {}) {
    try {
      const { limit = 10, offset = 0 } = options;
      const { count, rows } = await db.Company.findAndCountAll({
        limit,
        offset,
        order: [['id', 'ASC']]
      });
      return { total: count, companies: rows };
    } catch (error) {
      throw new Error(`Error fetching companies: ${error.message}`);
    }
  }

  /**
   * Get company by ID
   */
  static async findById(id) {
    try {
      const company = await db.Company.findByPk(id);
      if (!company) {
        throw new Error('Company not found');
      }
      return company;
    } catch (error) {
      throw new Error(`Error fetching company: ${error.message}`);
    }
  }

  /**
   * Get company by name
   */
  static async findByName(name) {
    try {
      const company = await db.Company.findOne({
        where: db.Sequelize.where(
          db.Sequelize.fn('LOWER', db.Sequelize.col('name')),
          db.Sequelize.Op.eq,
          name.toLowerCase()
        )
      });
      return company;
    } catch (error) {
      throw new Error(`Error fetching company by name: ${error.message}`);
    }
  }

  /**
   * Create new company
   */
  static async create(data) {
    try {
      const { name, company_percent, agent_percent } = data;

      if (!name) {
        throw new Error('Name is required');
      }

      const company = await db.Company.create({
        name,
        company_percent,
        agent_percent
      });

      return company;
    } catch (error) {
      throw new Error(`Error creating company: ${error.message}`);
    }
  }

  /**
   * Update company
   */
  static async update(id, data) {
    try {
      const company = await this.findById(id);
      const updated = await company.update(data);
      return updated;
    } catch (error) {
      throw new Error(`Error updating company: ${error.message}`);
    }
  }

  /**
   * Upsert company (update if exists, create if not)
   */
  static async upsert(data, where = {}) {
    try {
      const [company, created] = await db.Company.findOrCreate({
        where: where || {},
        defaults: data
      });

      if (!created) {
        await company.update(data);
      }

      return { company, created };
    } catch (error) {
      throw new Error(`Error upserting company: ${error.message}`);
    }
  }

  /**
   * Delete company
   */
  static async delete(id) {
    try {
      const company = await this.findById(id);
      await company.destroy();
      return { message: 'Company deleted successfully' };
    } catch (error) {
      throw new Error(`Error deleting company: ${error.message}`);
    }
  }

  /**
   * Search companies by name
   */
  static async search(query, options = {}) {
    try {
      const { limit = 10, offset = 0 } = options;
      const companies = await db.Company.findAll({
        where: {
          name: { [db.Sequelize.Op.iLike]: `%${query}%` }
        },
        limit,
        offset
      });
      return companies;
    } catch (error) {
      throw new Error(`Error searching companies: ${error.message}`);
    }
  }
}

module.exports = CompanyService;
