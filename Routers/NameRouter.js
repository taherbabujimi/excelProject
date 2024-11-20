const { addName } = require("../Controllers/NameController");

const { verifyJWT } = require("../Middlewares/authMiddleware");

const nameRoute = require("express").Router();

nameRoute.post("/addName", verifyJWT, addName);

module.exports = nameRoute;
