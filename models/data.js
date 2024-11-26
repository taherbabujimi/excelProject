"use strict";

const { Model } = require("sequelize");

module.exports = (sequelize, DataTypes) => {
  class Data extends Model {
    static associate(models) {
      // Data.hasOne(models.Name, { foreignKey: "id", as: "Name" });
      Data.belongsTo(models.Name, { foreignKey: "name", as: "Name" });

      this.belongsTo(models.User, {
        foreignKey: { field: "createdBy" },
        as: "CreatedBy",
      });
    }
  }
  Data.init(
    {
      id: {
        type: DataTypes.INTEGER,
        autoIncrement: true,
        primaryKey: true,
        allowNull: false,
      },
      name: {
        type: DataTypes.INTEGER,
        references: { model: "names", key: "id" },
        allowNull: false,
      },
      date: {
        type: DataTypes.DATE,
        allowNull: false,
      },
      amount: {
        type: DataTypes.FLOAT,
        allowNull: false,
      },
      createdBy: {
        type: DataTypes.INTEGER,
        allowNull: false,
      },
    },
    {
      sequelize,
      modelName: "Data",
      tableName: "datas",
      timestamps: true,
    }
  );

  return Data;
};
