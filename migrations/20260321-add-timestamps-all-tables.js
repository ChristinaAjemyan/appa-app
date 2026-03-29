'use strict';

module.exports = {
  async up(queryInterface, Sequelize) {
    // Add timestamps to insurance_policies table
    await queryInterface.addColumn('insurance_policies', 'createdAt', {
      type: Sequelize.DATE,
      allowNull: false,
      defaultValue: Sequelize.fn('NOW')
    });
    await queryInterface.addColumn('insurance_policies', 'updatedAt', {
      type: Sequelize.DATE,
      allowNull: false,
      defaultValue: Sequelize.fn('NOW')
    });

    // Add timestamps to agents table
    await queryInterface.addColumn('agents', 'createdAt', {
      type: Sequelize.DATE,
      allowNull: false,
      defaultValue: Sequelize.fn('NOW')
    });
    await queryInterface.addColumn('agents', 'updatedAt', {
      type: Sequelize.DATE,
      allowNull: false,
      defaultValue: Sequelize.fn('NOW')
    });

    // Add timestamps to agents_percentage table
    await queryInterface.addColumn('agents_percentage', 'createdAt', {
      type: Sequelize.DATE,
      allowNull: false,
      defaultValue: Sequelize.fn('NOW')
    });
    await queryInterface.addColumn('agents_percentage', 'updatedAt', {
      type: Sequelize.DATE,
      allowNull: false,
      defaultValue: Sequelize.fn('NOW')
    });

    // Add timestamps to company table
    await queryInterface.addColumn('company', 'createdAt', {
      type: Sequelize.DATE,
      allowNull: false,
      defaultValue: Sequelize.fn('NOW')
    });
    await queryInterface.addColumn('company', 'updatedAt', {
      type: Sequelize.DATE,
      allowNull: false,
      defaultValue: Sequelize.fn('NOW')
    });

    // Add timestamps to companies_percentage table
    await queryInterface.addColumn('companies_percentage', 'createdAt', {
      type: Sequelize.DATE,
      allowNull: false,
      defaultValue: Sequelize.fn('NOW')
    });
    await queryInterface.addColumn('companies_percentage', 'updatedAt', {
      type: Sequelize.DATE,
      allowNull: false,
      defaultValue: Sequelize.fn('NOW')
    });

    // Add timestamps to regions table
    await queryInterface.addColumn('regions', 'createdAt', {
      type: Sequelize.DATE,
      allowNull: false,
      defaultValue: Sequelize.fn('NOW')
    });
    await queryInterface.addColumn('regions', 'updatedAt', {
      type: Sequelize.DATE,
      allowNull: false,
      defaultValue: Sequelize.fn('NOW')
    });

    // Add timestamps to users table
    await queryInterface.addColumn('users', 'createdAt', {
      type: Sequelize.DATE,
      allowNull: false,
      defaultValue: Sequelize.fn('NOW')
    });
    await queryInterface.addColumn('users', 'updatedAt', {
      type: Sequelize.DATE,
      allowNull: false,
      defaultValue: Sequelize.fn('NOW')
    });
  },

  async down(queryInterface, Sequelize) {
    // Remove timestamps from all tables
    const tables = ['insurance_policies', 'agents', 'agents_percentage', 'company', 'companies_percentage', 'regions', 'users'];
    
    for (const table of tables) {
      await queryInterface.removeColumn(table, 'createdAt');
      await queryInterface.removeColumn(table, 'updatedAt');
    }
  }
};
