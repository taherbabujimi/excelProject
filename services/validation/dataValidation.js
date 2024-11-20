const joi = require("joi");

const { CATEGORIES } = require("../constants");
const categories = Object.values(CATEGORIES);
const { errorResponseWithoutData, errorResponseData } = require("../responses");
const { messages } = require("../messages");

module.exports = {
  addDataSchema(body, res) {
    const schema = joi.object({
      name: joi
        .string()
        .pattern(/^[a-zA-Z\s]+$/)
        .required()
        .min(2)
        .max(30),
      category: joi
        .string()
        .valid(...categories)
        .required(),
      date: joi.date().required(),
      amount: joi.number().required(),
    });

    const validationResult = schema.validate(body);

    let errorMessage = "";

    if (validationResult.error) {
      errorMessage = `${messages.errorWhileValidating}: ${validationResult.error}`;
      return errorMessage;
    } else {
      return false;
    }
  },

  getDataSchema(body, res) {
    const schema = joi.object({
      startDate: joi.date().required(),
      endDate: joi.date().required(),
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
