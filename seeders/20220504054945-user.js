"use strict";

module.exports = {
  async up (queryInterface) {
      await queryInterface.bulkInsert("users", [
          {
            name: "John Winchester",
            createdAt: new Date(),
            updatedAt: new Date(),
          },
          {
            name: "Sam Winchester",
            createdAt: new Date(),
            updatedAt: new Date(),
          }
     ], {});
  },
  async down (queryInterface) {
     await queryInterface.bulkDelete("users", null, {});
  }
};

