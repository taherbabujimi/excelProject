"use strict";

const { Model } = require("sequelize");

const { CATEGORIES } = require("../services/constants");
const categories = Object.values(CATEGORIES);

module.exports = (sequelize, DataTypes) => {
  class Name extends Model {}

  Name.init(
    {
      id: {
        type: DataTypes.INTEGER,
        allowNull: false,
        primaryKey: true,
        autoIncrement: true,
      },
      name: {
        type: DataTypes.STRING,
        allowNull: false,
      },
      userId: {
        type: DataTypes.INTEGER,
        references: { model: "users", key: "id" },
        allowNull: false,
      },
      category: {
        type: DataTypes.ENUM(categories),
        allowNull: false,
      },
    },
    {
      sequelize,
      modelName: "Name",
      tableName: "names",
      timestamps: true,
    }
  );

  return Name;
};
