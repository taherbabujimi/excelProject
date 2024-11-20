"use strict";

const bcrypt = require("bcrypt");
const jwt = require("jsonwebtoken");

module.exports = {
  async validPassword(password, user) {
    return bcrypt.compare(password, user.password);
  },

  async generateAccessToken(user) {
    return jwt.sign(
      {
        id: user.id,
        username: user.username,
        email: user.email,
      },
      process.env.ACCESS_TOKEN_SECRET,
      {
        expiresIn: process.env.ACCESS_TOKEN_EXPIRY,
      }
    );
  },
};
