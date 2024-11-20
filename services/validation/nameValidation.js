"use strict";

const joi = require("joi");

const { errorResponseData } = require("../responses");
const { messages } = require("../messages");

module.exports = {
  addNameSchema(body, res) {
    const schema = joi.object({
      name: joi
        .string()
        .pattern(/^[a-zA-Z]+$/)
        .required()
        .min(3)
        .max(30),
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
