const { db } = require('../database');

class CompanyPercentageService {
  /**
   * Get all company percentages with optional pagination
   */
  static async findAll(options = {}) {
    try {
      const { limit = 10, offset = 0 } = options;
      const records = await db.CompanyPercentage.findAll({
        limit,
        offset,
        order: [['id', 'ASC']]
      });
      return records;
    } catch (error) {
      throw new Error(`Error fetching company percentages: ${error.message}`);
    }
  }

  /**
   * Get all company percentages with count for pagination
   */
  static async findAllWithCount(options = {}) {
    try {
      const { limit = 10, offset = 0 } = options;
      const { count, rows } = await db.CompanyPercentage.findAndCountAll({
        limit,
        offset,
        order: [['id', 'ASC']]
      });
      return { total: count, records: rows };
    } catch (error) {
      throw new Error(`Error fetching company percentages: ${error.message}`);
    }
  }

  /**
   * Get company percentage by ID
   */
  static async findById(id) {
    try {
      const record = await db.CompanyPercentage.findByPk(id);
      if (!record) {
        throw new Error('Company percentage record not found');
      }
      return record;
    } catch (error) {
      throw new Error(`Error fetching company percentage: ${error.message}`);
    }
  }

  /**
   * Get percentages by company
   */
  static async findByCompany(company) {
    try {
      const records = await db.CompanyPercentage.findAll({
        where: { company: db.Sequelize.where(db.Sequelize.fn('LOWER', db.Sequelize.col('company')), 'LIKE', `%${company.toLowerCase()}%`) },
        order: [['id', 'ASC']]
      });
      return records;
    } catch (error) {
      throw new Error(`Error fetching percentages for company: ${error.message}`);
    }
  }

  /**
   * Create new company percentage record
   */
  static async create(data) {
    try {
      const { company, percentage } = data;

      if (!company || percentage === undefined) {
        throw new Error('Company and percentage are required');
      }

      const record = await db.CompanyPercentage.create(data);
      return record;
    } catch (error) {
      throw new Error(`Error creating company percentage: ${error.message}`);
    }
  }

  /**
   * Update company percentage record
   */
  static async update(id, data) {
    try {
      const record = await this.findById(id);
      const updated = await record.update(data);
      return updated;
    } catch (error) {
      throw new Error(`Error updating company percentage: ${error.message}`);
    }
  }

  /**
   * Upsert company percentage record (update if exists, create if not)
   */
  static async upsert(data, where = {}) {
    try {
      const [record, created] = await db.CompanyPercentage.findOrCreate({
        where: where || {},
        defaults: data
      });

      if (!created) {
        await record.update(data);
      }

      return { record, created };
    } catch (error) {
      throw new Error(`Error upserting company percentage: ${error.message}`);
    }
  }

  /**
   * Delete company percentage record
   */
  static async delete(id) {
    try {
      const record = await this.findById(id);
      await record.destroy();
      return { message: 'Company percentage record deleted successfully' };
    } catch (error) {
      throw new Error(`Error deleting company percentage: ${error.message}`);
    }
  }

  /**
   * Find matching exception for given criteria
   */
  static async findMatchingException(company, criteria = {}) {
    try {
      const {
        agent_code_in,
        region_in,
        bm_exact,
        bm_min,
        bm_max,
        hp_min,
        hp_max,
        brand_in,
        period
      } = criteria;

      const where = {
        company: db.Sequelize.where(db.Sequelize.fn('LOWER', db.Sequelize.col('company')), db.Sequelize.Op.eq, company.toLowerCase())
      };

      if (agent_code_in) {
        where.agent_code_in = { [db.Sequelize.Op.iLike]: `%${agent_code_in}%` };
      }

      if (region_in) {
        where.region_in = { [db.Sequelize.Op.iLike]: `%${region_in}%` };
      }

      if (bm_exact) {
        where.bm_exact = bm_exact;
      }

      if (brand_in) {
        where.brand_in = { [db.Sequelize.Op.iLike]: `%${brand_in}%` };
      }

      if (period) {
        where.period = db.Sequelize.where(db.Sequelize.fn('LOWER', db.Sequelize.col('period')), db.Sequelize.Op.eq, period.toLowerCase());
      }

      const record = await db.CompanyPercentage.findOne({
        where,
        limit: 1
      });

      return record;
    } catch (error) {
      throw new Error(`Error finding matching exception: ${error.message}`);
    }
  }
}

module.exports = CompanyPercentageService;
