const IndexRoute = require("express").Router();
const userRoute = require("../Routers/UserRouter");
const nameRoute = require("../Routers/NameRouter");
const dataRoute = require("../Routers/DataRoute");
const cronJobs = require("../Cron_Jobs/sendData");

IndexRoute.use("/v1/users", userRoute);
IndexRoute.use("/v1/names", nameRoute);
IndexRoute.use("/v1/data", dataRoute);

module.exports = IndexRoute;
