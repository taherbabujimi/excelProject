"use strict";

const { CATEGORIES } = require("../services/constants");
const categories = Object.values(CATEGORIES);

module.exports = {
  async up(queryInterface, Sequelize) {
    await queryInterface.addColumn("names", "category", {
      type: Sequelize.ENUM(categories),
      allowNull: false,
    });
  },

  async down(queryInterface, Sequelize) {
    await queryInterface.removeColumn("names", "category");
  },
};
