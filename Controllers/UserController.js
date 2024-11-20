const Models = require("../models/index");

const {
  registerUserSchema,
  loginUserSchema,
} = require("../services/validation/userValidation");

const { validPassword, generateAccessToken } = require("../services/helpers")

const { messages } = require("../services/messages");

const {
  successResponseData,
  errorResponseWithoutData,
} = require("../services/responses");

module.exports.registerUser = async (req, res) => {
  try {
    const validationResponse = registerUserSchema(req.body, res);
    if (validationResponse !== false) return;

    let { username, password } = req.body;
    username = username.replace(/ +/gm, "");

    const email = req.body.email.toLowerCase();

    const oldUser = await Models.User.findOne({
      where: { email: email },
    });

    if (oldUser) {
      return errorResponseWithoutData(res, messages.userAlreadyExists, 400);
    }

    const user = await Models.User.create({
      username,
      email,
      password,
    });

    return successResponseData(
      res,
      user,
      200,
      messages.userCreatedSuccessfully
    );
  } catch (error) {
    return errorResponseWithoutData(
      res,
      `${messages.somethingWentWrong} : ${error}`,
      400
    );
  }
};

module.exports.userLogin = async (req, res) => {
  try {
    const validationResponse = loginUserSchema(req.body, res);

    if (validationResponse !== false) return;

    const { email, password } = req.body;

    const user = await Models.User.findOne({
      where: { email: email },
    });

    if (!user) {
      return errorResponseWithoutData(res, messages.userNotExists, 400);
    }

    const isPasswordValid = await validPassword(password, user);

    if (!isPasswordValid) {
      return errorResponseWithoutData(res, messages.incorrectCredentials, 400);
    }

    const accessToken = await generateAccessToken(user);

    const userData = {
      id: user.id,
      username: user.username,
      email: user.email,
      usertype: user.usertype,
    };

    return successResponseData(res, userData, 200, messages.userLoginSuccess, {
      token: accessToken,
    });
  } catch (error) {
    console.log("error: ",error)
    return errorResponseWithoutData(
      res,
      `${messages.somethingWentWrong}: ${error}`
    );
  }
};
