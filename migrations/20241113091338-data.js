"use strict";

const { CATEGORIES } = require("../services/constants");
const categories = Object.values(CATEGORIES);

module.exports = {
  async up(queryInterface, Sequelize) {
    await queryInterface.removeColumn("datas", "category");
  },

  async down(queryInterface, Sequelize) {
    await queryInterface.addColumn("datas", "category", {
      type: Sequelize.ENUM(categories),
      allowNull: false,
    });
  },
};
