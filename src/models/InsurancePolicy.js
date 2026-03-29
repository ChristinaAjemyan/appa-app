const { DataTypes } = require('sequelize');

module.exports = (sequelize) => {
  const InsurancePolicy = sequelize.define(
    'insurance_policies',
    {
      id: {
        type: DataTypes.INTEGER,
        primaryKey: true,
        autoIncrement: true
      },
      company: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      agent_company_code: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      agent_inner_code: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      agent_name: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      polis_number: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      owner_name: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      start_date: {
        type: DataTypes.DATE,
        allowNull: true
      },
      end_date: {
        type: DataTypes.DATE,
        allowNull: true
      },
      region: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      phone_number: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      bm_class: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      car_model: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      car_number: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      hp: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      period: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      info: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      price: {
        type: DataTypes.REAL,
        allowNull: true
      },
      agent_percent: {
        type: DataTypes.REAL,
        allowNull: true
      },
      company_percent: {
        type: DataTypes.REAL,
        allowNull: true
      },
      agent_income: {
        type: DataTypes.REAL,
        allowNull: true
      },
      income: {
        type: DataTypes.REAL,
        allowNull: true
      }
    },
    {
      tableName: 'insurance_policies',
      freezeTableName: true,
      timestamps: false
    }
  );

  return InsurancePolicy;
};
