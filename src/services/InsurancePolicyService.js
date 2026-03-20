const { db } = require('../database');

class InsurancePolicyService {
  /**
   * Get all insurance policies with optional pagination and sorting
   */
  static async findAll(options = {}) {
    try {
      const { limit = 10, offset = 0, sortBy = 'id', sortOrder = 'ASC' } = options;
      
      const allowedColumns = ['id', 'company', 'agent_inner_code', 'polis_number', 'owner_name',
        'start_date', 'end_date', 'region', 'price', 'agent_income', 'income'];
      const validSortBy = allowedColumns.includes(sortBy) ? sortBy : 'id';
      const validSortOrder = sortOrder.toUpperCase() === 'DESC' ? 'DESC' : 'ASC';

      const policies = await db.InsurancePolicy.findAll({
        limit,
        offset,
        order: [[validSortBy, validSortOrder]]
      });
      return policies;
    } catch (error) {
      throw new Error(`Error fetching insurance policies: ${error.message}`);
    }
  }

  /**
   * Get all insurance policies with count for pagination
   */
  static async findAllWithCount(options = {}) {
    try {
      const { limit = 10, offset = 0, sortBy = 'id', sortOrder = 'ASC' } = options;

      const allowedColumns = ['id', 'company', 'agent_inner_code', 'polis_number', 'owner_name',
        'start_date', 'end_date', 'region', 'price', 'agent_income', 'income'];
      const validSortBy = allowedColumns.includes(sortBy) ? sortBy : 'id';
      const validSortOrder = sortOrder.toUpperCase() === 'DESC' ? 'DESC' : 'ASC';

      const { count, rows } = await db.InsurancePolicy.findAndCountAll({
        limit,
        offset,
        order: [[validSortBy, validSortOrder]]
      });
      return { total: count, policies: rows };
    } catch (error) {
      throw new Error(`Error fetching insurance policies: ${error.message}`);
    }
  }

  /**
   * Get insurance policy by ID
   */
  static async findById(id) {
    try {
      const policy = await db.InsurancePolicy.findByPk(id);
      if (!policy) {
        throw new Error('Insurance policy not found');
      }
      return policy;
    } catch (error) {
      throw new Error(`Error fetching insurance policy: ${error.message}`);
    }
  }

  /**
   * Search insurance policies with filtering
   */
  static async search(filters = {}, options = {}) {
    try {
      const { limit = 10, offset = 0, sortBy = 'id', sortOrder = 'ASC' } = options;
      const where = {};

      // Build where clause from filters
      if (filters.company) {
        where.company = { [db.Sequelize.Op.iLike]: `%${filters.company}%` };
      }
      if (filters.agent_inner_code) {
        where.agent_inner_code = { [db.Sequelize.Op.iLike]: `%${filters.agent_inner_code}%` };
      }
      if (filters.polis_number) {
        where.polis_number = { [db.Sequelize.Op.iLike]: `%${filters.polis_number}%` };
      }
      if (filters.region) {
        where.region = { [db.Sequelize.Op.iLike]: `%${filters.region}%` };
      }
      if (filters.owner_name) {
        where.owner_name = { [db.Sequelize.Op.iLike]: `%${filters.owner_name}%` };
      }
      if (filters.minPrice !== undefined) {
        where.price = { [db.Sequelize.Op.gte]: parseFloat(filters.minPrice) };
      }
      if (filters.maxPrice !== undefined) {
        if (where.price) {
          where.price[db.Sequelize.Op.lte] = parseFloat(filters.maxPrice);
        } else {
          where.price = { [db.Sequelize.Op.lte]: parseFloat(filters.maxPrice) };
        }
      }
      if (filters.startDate) {
        where.start_date = { [db.Sequelize.Op.gte]: filters.startDate };
      }
      if (filters.endDate) {
        if (where.start_date) {
          where.start_date[db.Sequelize.Op.lte] = filters.endDate;
        } else {
          where.start_date = { [db.Sequelize.Op.lte]: filters.endDate };
        }
      }

      const allowedColumns = ['id', 'company', 'agent_inner_code', 'polis_number', 'owner_name',
        'start_date', 'end_date', 'region', 'price', 'agent_income', 'income'];
      const validSortBy = allowedColumns.includes(sortBy) ? sortBy : 'id';
      const validSortOrder = sortOrder.toUpperCase() === 'DESC' ? 'DESC' : 'ASC';

      const { count, rows } = await db.InsurancePolicy.findAndCountAll({
        where,
        limit,
        offset,
        order: [[validSortBy, validSortOrder]]
      });

      return { total: count, policies: rows };
    } catch (error) {
      throw new Error(`Error searching insurance policies: ${error.message}`);
    }
  }

  /**
   * Create new insurance policy
   */
  static async create(data) {
    try {
      const { company, polis_number } = data;

      if (!company || !polis_number) {
        throw new Error('Company and polis_number are required');
      }

      const policy = await db.InsurancePolicy.create(data);
      return policy;
    } catch (error) {
      throw new Error(`Error creating insurance policy: ${error.message}`);
    }
  }

  /**
   * Create multiple insurance policies
   */
  static async createBulk(dataArray) {
    try {
      const policies = await db.InsurancePolicy.bulkCreate(dataArray);
      return { created: policies.length, policies };
    } catch (error) {
      throw new Error(`Error creating insurance policies: ${error.message}`);
    }
  }

  /**
   * Update insurance policy
   */
  static async update(id, data) {
    try {
      const policy = await this.findById(id);
      const updated = await policy.update(data);
      return updated;
    } catch (error) {
      throw new Error(`Error updating insurance policy: ${error.message}`);
    }
  }

  /**
   * Upsert insurance policy (update if exists, create if not)
   */
  static async upsert(data, where = {}) {
    try {
      const [policy, created] = await db.InsurancePolicy.findOrCreate({
        where: where || {},
        defaults: data
      });

      if (!created) {
        await policy.update(data);
      }

      return { policy, created };
    } catch (error) {
      throw new Error(`Error upserting insurance policy: ${error.message}`);
    }
  }

  /**
   * Delete insurance policy
   */
  static async delete(id) {
    try {
      const policy = await this.findById(id);
      await policy.destroy();
      return { message: 'Insurance policy deleted successfully' };
    } catch (error) {
      throw new Error(`Error deleting insurance policy: ${error.message}`);
    }
  }

  /**
   * Get policies by company
   */
  static async findByCompany(company, options = {}) {
    try {
      const { limit = 10, offset = 0 } = options;
      const policies = await db.InsurancePolicy.findAll({
        where: {
          company: db.Sequelize.where(
            db.Sequelize.fn('LOWER', db.Sequelize.col('company')),
            db.Sequelize.Op.eq,
            company.toLowerCase()
          )
        },
        limit,
        offset,
        order: [['id', 'ASC']]
      });
      return policies;
    } catch (error) {
      throw new Error(`Error fetching policies by company: ${error.message}`);
    }
  }

  /**
   * Get policies by agent
   */
  static async findByAgent(agentCode, options = {}) {
    try {
      const { limit = 10, offset = 0 } = options;
      const policies = await db.InsurancePolicy.findAll({
        where: {
          agent_inner_code: { [db.Sequelize.Op.iLike]: `%${agentCode}%` }
        },
        limit,
        offset
      });
      return policies;
    } catch (error) {
      throw new Error(`Error fetching policies by agent: ${error.message}`);
    }
  }

  /**
   * Get aggregate statistics for policies
   */
  static async getStatistics(filters = {}) {
    try {
      const where = {};

      if (filters.company) {
        where.company = { [db.Sequelize.Op.iLike]: `%${filters.company}%` };
      }

      const stats = await db.InsurancePolicy.findAll({
        attributes: [
          [db.Sequelize.fn('COUNT', db.Sequelize.col('id')), 'total_policies'],
          [db.Sequelize.fn('SUM', db.Sequelize.col('price')), 'total_price'],
          [db.Sequelize.fn('AVG', db.Sequelize.col('price')), 'avg_price'],
          [db.Sequelize.fn('SUM', db.Sequelize.col('agent_income')), 'total_agent_income'],
          [db.Sequelize.fn('SUM', db.Sequelize.col('income')), 'total_income']
        ],
        where,
        raw: true
      });

      return stats[0];
    } catch (error) {
      throw new Error(`Error getting statistics: ${error.message}`);
    }
  }
}

module.exports = InsurancePolicyService;
