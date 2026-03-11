const { db } = require('../database');

class AgentService {
  /**
   * Get all agents with optional pagination
   */
  static async findAll() {
    try {
      const agents = await db.Agent.findAll({
        order: [['id', 'ASC']]
      });
      return agents;
    } catch (error) {
      throw new Error(`Error fetching agents: ${error.message}`);
    }
  }

  /**
   * Get all agents with count for pagination
   */
  static async findAllWithCount(options = {}) {
    try {
      const { limit = 10, offset = 0 } = options;
      const { count, rows } = await db.Agent.findAndCountAll({
        limit,
        offset,
        order: [['id', 'ASC']]
      });
      console.log(`Total agents: ${count}, Fetched agents: ${rows.length}`);
      return { total: count, agents: rows };
    } catch (error) {
      throw new Error(`Error fetching agents: ${error.message}`);
    }
  }

  /**
   * Get agent by ID
   */
  static async findById(id) {
    try {
      const agent = await db.Agent.findByPk(id);
      if (!agent) {
        throw new Error('Agent not found');
      }
      return agent;
    } catch (error) {
      throw new Error(`Error fetching agent: ${error.message}`);
    }
  }

  /**
   * Create new agent
   */
  static async create(data) {
    try {
      const { code, name, phone_number, NAIRI, REGO, ARMENIA, SIL, INGO, LIGA } = data;

      if (!code || !name) {
        throw new Error('Code and name are required');
      }

      const agent = await db.Agent.create({
        code,
        name,
        phone_number,
        NAIRI,
        REGO,
        ARMENIA,
        SIL,
        INGO,
        LIGA
      });

      return agent;
    } catch (error) {
      throw new Error(`Error creating agent: ${error.message}`);
    }
  }

  /**
   * Update agent
   */
  static async update(id, data) {
    try {
      const agent = await this.findById(id);
      const updated = await agent.update(data);
      return updated;
    } catch (error) {
      throw new Error(`Error updating agent: ${error.message}`);
    }
  }

  /**
   * Upsert agent (update if exists, create if not)
   */
  static async upsert(data, where = {}) {
    try {
      const [agent, created] = await db.Agent.findOrCreate({
        where: where || {},
        defaults: data
      });

      if (!created) {
        await agent.update(data);
      }
      return { agent, created };
    } catch (error) {
      throw new Error(`Error upserting agent: ${error.message}`);
    }
  }

  /**
   * Delete agent
   */
  static async delete(id) {
    try {
      const agent = await this.findById(id);
      await agent.destroy();
      return { message: 'Agent deleted successfully' };
    } catch (error) {
      throw new Error(`Error deleting agent: ${error.message}`);
    }
  }

  /**
   * Search agents by name or code
   */
  static async search(query, options = {}) {
    try {
      const { limit = 10, offset = 0 } = options;
      const agents = await db.Agent.findAll({
        where: {
          [db.Sequelize.Op.or]: [
            { code: { [db.Sequelize.Op.iLike]: `%${query}%` } },
            { name: { [db.Sequelize.Op.iLike]: `%${query}%` } }
          ]
        },
        limit,
        offset
      });
      return agents;
    } catch (error) {
      throw new Error(`Error searching agents: ${error.message}`);
    }
  }
}

module.exports = AgentService;
