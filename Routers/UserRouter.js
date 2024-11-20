const { registerUser, userLogin } = require("../Controllers/UserController.js");

const userRoute = require("express").Router();

userRoute.post("/registerUser", registerUser);
userRoute.post("/userLogin", userLogin);

module.exports = userRoute;
