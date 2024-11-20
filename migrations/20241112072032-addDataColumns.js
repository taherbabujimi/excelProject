"use strict";

module.exports = {
  async up(queryInterface, Sequelize) {
    await queryInterface.addColumn("datas", "createdAt", {
      type: Sequelize.DATE,
      allowNull: false,
    });
    await queryInterface.addColumn("datas", "updatedAt", {
      type: Sequelize.DATE,
      allowNull: false,
    });
  },

  async down(queryInterface, Sequelize) {
    await queryInterface.removeColumn("datas", "createdAt");
    await queryInterface.removeColumn("datas", "updatedAt");
  },
};
