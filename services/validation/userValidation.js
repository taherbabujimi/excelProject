"use strict";

const joi = require("joi");
const { errorResponseData } = require("../responses");
const { messages } = require("../messages");

module.exports = {
  registerUserSchema(body, res) {
    const schema = joi.object({
      username: joi.string().min(3).max(30).required(),
      email: joi.string().email(),
      password: joi.string().required().min(3).max(30),
    });

    const validationResult = schema.validate(body);

    if (validationResult.error) {
      return errorResponseData(
        res,
        messages.errorWhileValidating,
        validationResult.error.details
      );
    } else {
      return false;
    }
  },

  loginUserSchema(body, res) {
    const schema = joi.object({
      email: joi.string().email(),
      password: joi.string().required().min(3).max(30),
    });

    const validationResult = schema.validate(body);

    if (validationResult.error) {
      return errorResponseData(
        res,
        messages.errorWhileValidating,
        validationResult.error.details
      );
    } else {
      return false;
    }
  },
};
