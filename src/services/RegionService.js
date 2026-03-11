const { db } = require('../database');

class RegionService {
  /**
   * Find all regions
   */
  static async findAll() {
    return await db.Region.findAll({
      order: [['region_code', 'ASC']]
    });
  }

  /**
   * Find region by ID
   */
  static async findById(id) {
    const region = await db.Region.findByPk(id);
    if (!region) {
      throw new Error('Region not found');
    }
    return region;
  }

  /**
   * Find region by code
   */
  static async findByCode(regionCode) {
    return await db.Region.findOne({
      where: { region_code: regionCode }
    });
  }

  /**
   * Create new region
   */
  static async create(data) {
    const { region_code, name } = data;

    if (!region_code || !name) {
      throw new Error('region_code and name are required');
    }

    return await db.Region.create({
      region_code,
      name
    });
  }

  /**
   * Update region
   */
  static async update(id, data) {
    const region = await this.findById(id);
    return await region.update(data);
  }

  /**
   * Delete region
   */
  static async delete(id) {
    const region = await this.findById(id);
    await region.destroy();
    return { message: 'Region deleted successfully' };
  }
}

module.exports = RegionService;
