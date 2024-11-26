const {
  addData,
  getData,
  addFormula,
} = require("../Controllers/DataController");

const { verifyJWT } = require("../Middlewares/authMiddleware");

const dataRoute = require("express").Router();
const multer = require("multer");
const upload = multer({ storage: multer.memoryStorage() });

dataRoute.post("/addData", verifyJWT, upload.single("file"), addData);

dataRoute.get("/getData", verifyJWT, getData);

dataRoute.post("/addFormula", verifyJWT, addFormula);

module.exports = dataRoute;
