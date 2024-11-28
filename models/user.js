"use strict";
const { Model } = require("sequelize");
const { USER_TYPE } = require("../services/constants");
const USERTYPE = Object.values(USER_TYPE);
const bcrypt = require("bcrypt");
const jwt = require("jsonwebtoken");

module.exports = (sequelize, DataTypes) => {
  class User extends Model {
    static associate(models) {
      this.hasMany(models.Formula, {
        foreignKey: { field: "createdBy" },
        as: "Formula",
      });

      this.hasMany(models.Data, {
        foreignKey: { field: "createdBy" },
        as: "Data",
      });

      // this.hasMany(models.Name, {
      //   foreignKey: { field: "userId" },
      //   as: "Name",
      // });
    }
  }
  User.init(
    {
      id: {
        type: DataTypes.INTEGER,
        allowNull: false,
        primaryKey: true,
        autoIncrement: true,
      },
      username: {
        type: DataTypes.STRING,
        allowNull: false,
      },
      email: {
        type: DataTypes.STRING,
        allowNull: false,
        unique: true,
      },
      password: {
        type: DataTypes.STRING,
        allowNull: false,
      },
      usertype: {
        type: DataTypes.ENUM(USERTYPE),
        defaultValue: USER_TYPE.USER,
        allowNull: false,
      },
    },
    {
      sequelize,
      modelName: "User",
      tableName: "users",
      timestamps: true,
    }
  );

  User.beforeCreate(async (user, options) => {
    return await bcrypt
      .hash(user.password, 10)
      .then((hash) => {
        user.password = hash;
      })
      .catch((error) => {
        throw new Error(error);
      });
  });

  User.prototype.generateHash = async (password) => {
    return await bcrypt.hashSync(password, bcrypt.genSaltSync(10), null);
  };

  User.prototype.validPassword = async (password) => {
    return await bcrypt.compare(password, this.password);
  };

  User.prototype.generateAccessToken = function () {
    return jwt.sign({
      id: this.id,
      username: this.username,
      email: this.email,
    });
  };

  return User;
};
