const { DataTypes } = require('sequelize');

module.exports = (sequelize) => {
  const Agent = sequelize.define(
    'agents',
    {
      id: {
        type: DataTypes.INTEGER,
        primaryKey: true,
        autoIncrement: true
      },
      code: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      name: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      phone_number: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      nairi: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      rego: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      armenia: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      sil: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      ingo: {
        type: DataTypes.TEXT,
        allowNull: true
      },
      liga: {
        type: DataTypes.TEXT,
        allowNull: true
      }
    },
    {
      tableName: 'agents',
      freezeTableName: true,
      timestamps: true,
      createdAt: 'createdAt',
      updatedAt: 'updatedAt'
    }
  );

  return Agent;
};
