const { DataTypes } = require('sequelize');

module.exports = (sequelize) => {
  const CompanyPercentage = sequelize.define(
    'companies_percentage',
    {
      id: {
        type: DataTypes.INTEGER,
        primaryKey: true,
        autoIncrement: true
      },
      company: {
        type: DataTypes.TEXT,
        allowNull: false
      },
      agent_code_in: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      agent_code_not: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      region_in: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      region_not: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      bm_min: {
        type: DataTypes.INTEGER,
        allowNull: true
      },
      bm_max: {
        type: DataTypes.INTEGER,
        allowNull: true
      },
      bm_exact: {
        type: DataTypes.INTEGER,
        allowNull: true
      },
      brand_in: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      hp_min: {
        type: DataTypes.INTEGER,
        allowNull: true
      },
      hp_max: {
        type: DataTypes.INTEGER,
        allowNull: true
      },
      period: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      percentage: {
        type: DataTypes.REAL,
        allowNull: false
      }
    },
    {
      tableName: 'companies_percentage',
      freezeTableName: true,
      timestamps: true,
      createdAt: 'createdAt',
      updatedAt: 'updatedAt'
    }
  );

  return CompanyPercentage;
};
