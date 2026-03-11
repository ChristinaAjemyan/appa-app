const { DataTypes } = require('sequelize');

module.exports = (sequelize) => {
  const Region = sequelize.define(
    'regions',
    {
      id: {
        type: DataTypes.INTEGER,
        primaryKey: true,
        autoIncrement: true
      },
      region_code: {
        type: DataTypes.TEXT,
        allowNull: false,
        unique: true
      },
      name: {
        type: DataTypes.TEXT,
        allowNull: false
      }
    },
    {
      tableName: 'regions',
      timestamps: false
    }
  );

  return Region;
};
