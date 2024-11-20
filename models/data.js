"use strict";

const { Model } = require("sequelize");

module.exports = (sequelize, DataTypes) => {
  class Data extends Model {}
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
