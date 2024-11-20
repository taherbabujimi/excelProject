"use strict";

module.exports = {
  async up(queryInterface, Sequelize) {
    await queryInterface.changeColumn("datas", "amount", {
      type: Sequelize.FLOAT,
      allowNull: false,
    });
  },

  async down(queryInterface, Sequelize) {
    await queryInterface.changeColumn("datas", "amount", {
      type: Sequelize.INTEGER,
      allowNull: false,
    });
  },
};
