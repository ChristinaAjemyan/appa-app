const { DataTypes } = require('sequelize');

module.exports = (sequelize) => {
  const Company = sequelize.define(
    'company',
    {
      id: {
        type: DataTypes.INTEGER,
        primaryKey: true,
        autoIncrement: true
      },
      name: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      company_percent: {
        type: DataTypes.REAL,
        allowNull: true
      },
      agent_percent: {
        type: DataTypes.REAL,
        allowNull: true
      }
    },
    {
      tableName: 'company',
      freezeTableName: true,
      timestamps: true,
      createdAt: 'createdAt',
      updatedAt: 'updatedAt'
    }
  );

  return Company;
};
