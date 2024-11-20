"use strict";

const jwt = require("jsonwebtoken");
const {
  errorResponseWithoutData,
  errorResponseData,
} = require("../services/responses");
const Models = require("../models/index");
const { messages } = require("../services/messages");

module.exports = {
  async verifyJWT(req, res, next) {
    try {
      const token = req.header("Authorization").replace("Bearer ", "");

      if (!token) {
        return errorResponseWithoutData(res, messages.badRequest, 400);
      }

      const decodedToken = jwt.verify(token, process.env.ACCESS_TOKEN_SECRET);

      const user = await Models.User.findByPk(decodedToken.id);

      if (!user) {
        return errorResponseWithoutData(res, messages.invalidToken, 400);
      }

      req.user = user;

      next();
    } catch (error) {
      return errorResponseData(res, `${messages.invalidToken}: ${error}`, 400);
    }
  },
};
