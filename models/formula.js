"use strict";

const { Model } = require("sequelize");

module.exports = (sequelize, Datatypes) => {
  class Formula extends Model {
    static associate(models) {
      this.belongsTo(models.User, {
        foreignKey: { field: "createdBy" },
        as: "CreatedBy",
      });
    }
  }

  Formula.init(
    {
      id: {
        type: Datatypes.INTEGER,
        primaryKey: true,
        autoIncrement: true,
        allowNull: false,
      },
      formula: {
        type: Datatypes.STRING,
        allowNull: false,
      },
      createdBy: {
        type: Datatypes.INTEGER,
        references: { model: "users", key: "id" },
        allowNull: false,
      },
    },
    {
      sequelize,
      modelName: "Formula",
      tableName: "formulas",
      timestamps: true,
    }
  );

  return Formula;
};
