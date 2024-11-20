const Models = require("../models/index");

const { addNameSchema } = require("../services/validation/nameValidation");

const { Op } = require("sequelize");

const {
  errorResponseWithoutData,
  successResponseData,
  errorResponseData,
} = require("../services/responses");
const { messages } = require("../services/messages");

module.exports.addName = async (req, res) => {
  try {
    const validationResponse = addNameSchema(req.body, res);
    if (validationResponse !== false) return;

    const { name } = req.body;

    const existingName = await Models.Name.findOne({
      where: {
        [Op.and]: [{ name: name }, { userId: req.user.id }],
      },
    });

    if (existingName) {
      return errorResponseWithoutData(res, messages.nameAlreadyExists, 400);
    }

    const namee = await Models.Name.create({ name, userId: req.user.id });

    if (!namee) {
      return errorResponseWithoutData(res, messages.somethingWentWrong, 400);
    }

    return successResponseData(res, namee, 200, messages.nameAddedSuccess);
  } catch (error) {
    return errorResponseData(
      res,
      `${messages.somethingWentWrong}: ${error}`,
      400
    );
  }
};
