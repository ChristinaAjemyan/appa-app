const { DataTypes } = require('sequelize');

module.exports = (sequelize) => {
  const CalcResultStorage = sequelize.define(
    'calc_result_storage',
    {
      key: {
        type: DataTypes.TEXT,
        primaryKey: true,
        allowNull: false
      },
      value: {
        type: DataTypes.JSONB,
        allowNull: true
      }
    },
    {
      tableName: 'calc_result_storage',
      timestamps: false
    }
  );

  return CalcResultStorage;
};
