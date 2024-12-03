const Models = require("../models/index");
const { messages } = require("../services/messages");
const xlsx = require("xlsx");
const { Op } = require("sequelize");
const cron = require("node-cron");
const { sequelize } = require("../models/index");

async function calculateFormulas(userId) {
  try {
    const user = await Models.User.findOne({
      where: {
        id: userId,
      },
      include: [
        {
          model: Models.Formula,
          as: "Formula",
        },
      ],
    });

    console.log("User: ", JSON.stringify(user));

    let userName = user.dataValues.username;

    if (user.dataValues.Formula.length === 0) {
      console.log("no data");
      return false;
    }

    let excelData = [];

    let existedNames = [];

    for (const formulaItem of user.dataValues.Formula) {
      let formula = formulaItem.dataValues.formula;
      let formulaName = formulaItem.dataValues.formulaName;

      const pairs = formula.match(
        /\b\w+:(\d{4}-\d{2}-\d{2}|\d{4}-\d{2}|\d{4})\b/g
      );

      console.log("pairs: ", pairs);

      const result = pairs.reduce((acc, pair) => {
        const [key, value] = pair.split(":");
        const date = value.split("-");
        if (date.length === 3) {
          acc[key] = acc[key] || [];
          acc[key].push(`${date[0]}-${date[1]}-${date[2]}`);
        } else if (date.length === 2) {
          acc[key] = acc[key] || [];
          acc[key].push(`${date[0]}-${date[1]}`);
        } else {
          acc[key] = acc[key] || [];
          acc[key].push(date[0]);
        }
        return acc;
      }, {});

      console.log("result: ", result);

      const synonyms = Object.keys(result);
      console.log("synonyms: ", synonyms);

      let j = 0;
      let existedNamesResult = [];
      for (const item of synonyms) {
        console.log("item: ", item);
        console.log(result[item].length);

        for (let i = 0; i < result[item].length; i++) {
          let date = result[item][i].split("-");

          let firstDate;
          let lastDate;
          if (date.length === 2) {
            firstDate = new Date(result[item][i]);
            lastDate = new Date(
              firstDate.getFullYear(),
              firstDate.getMonth() + 1,
              0
            );
            lastDate = new Date(
              lastDate.getTime() +
                Math.abs(lastDate.getTimezoneOffset() * 60000)
            );
          } else if (date.length === 1) {
            firstDate = new Date(result[item][i]);
            lastDate = new Date(firstDate.getFullYear(), 11, 31);
            lastDate = new Date(
              lastDate.getTime() +
                Math.abs(lastDate.getTimezoneOffset() * 60000)
            );
          }

          existedNames.push(
            Models.Name.findOne({
              where: {
                [Op.and]: [{ synonym: item }, { userId: userId }],
              },
              include: [
                {
                  model: Models.Data,
                  as: "Data",
                  where: {
                    [Op.and]: [
                      {
                        date: {
                          [Op.gte]: firstDate,
                          [Op.lte]: lastDate,
                        },
                      },
                      {
                        createdBy: userId,
                      },
                    ],
                  },
                  attributes: [
                    [
                      sequelize.fn("SUM", sequelize.col("amount")),
                      "totalAmount",
                    ],
                  ],

                  required: false,
                },
              ],
            })
          );
        }
        await Promise.all(existedNames)
          .then((result) => {
            console.log(" result: ", JSON.stringify(result));
            existedNamesResult = result;
          })
          .catch((error) => {
            console.log(error);
            return errorResponseWithoutData(
              res,
              "Something went wrong while fetching the data.",
              400
            );
          });
      }

      let totalArray = [];

      for (const item of existedNamesResult) {
        if (item.id === null) {
          console.log("Item is null");
          totalArray.push(null);
          continue;
        }

        let itemTotal;

        if (item.Data.length === 0) {
          itemTotal = 0;
          // continue;
        } else {
          itemTotal = item.Data[0].dataValues.totalAmount;
        }

        totalArray.push(itemTotal);
      }

      const nameDatePatterns = formula.match(
        /(\w+):(\d{4}-\d{2}-\d{2}|\d{4}-\d{2}|\d{4})/g
      );

      if (totalArray.includes(null)) {
        excelData.push({
          [formulaName]: "Data For this formula not found",
        });
        continue;
      }

      const replacedFormula = nameDatePatterns.reduce((acc, pattern) => {
        const [name, date] = pattern.split(":");
        const index = totalArray.findIndex(
          (value, index) => index === nameDatePatterns.indexOf(pattern)
        );
        if (index !== -1) {
          return acc.replace(pattern, totalArray[index]);
        } else {
          return acc;
        }
      }, formula);

      const calculatedResult = eval(replacedFormula);

      excelData.push({ [formulaName]: calculatedResult.toFixed(2) });

      existedNames = [];
    }

    const excelDataRows = excelData.map((data) => {
      const formulaName = Object.keys(data)[0];
      const value = data[formulaName];
      return [formulaName, value];
    });

    return excelDataRows;
  } catch (error) {
    console.log(
      `${messages.somethingWentWrong} in sendData function : ${error}`
    );
  }
}

//*/3 * * * * *

//1 10 1 * *

cron.schedule("1 10 1 * *", async () => {
  try {
    function formatFirstDate(date, format) {
      const map = {
        mm: date.getMonth(),
        dd: date.getDate(),
        yy: date.getFullYear(),
      };

      return format.replace(/mm|dd|yy|yyy/gi, (matched) => map[matched]);
    }

    function formatLastDate(date, format) {
      const map = {
        mm: date.getMonth(),
        dd: date.getDate(),
        yy: date.getFullYear(),
      };

      map.mm += 1;

      return format.replace(/mm|dd|yy|yyy/gi, (matched) => map[matched]);
    }

    function GFG_Fun() {
      let date = new Date(),
        y = date.getFullYear(),
        m = date.getMonth();

      let firstDay = new Date(y, m, 1);
      let lastDay = new Date(y, m, 0);

      return { firstDay, lastDay };
    }

    const { firstDay, lastDay } = GFG_Fun();

    let formattedFirstDay = formatFirstDate(firstDay, "yy-mm-dd");
    let formattedLastDay = formatLastDate(lastDay, "yy-mm-dd");

    const users = await Models.User.count({});

    let offset = users / 10;
    offset = Math.ceil(offset);

    let userData = [];
    let promiseUserData = [];
    for (let i = 0; i < offset; i++) {
      let limit = 10;
      userData.push(
        Models.User.findAll({
          limit,
          offset: i * limit,
          include: [
            {
              model: Models.Data,
              as: "Data",
              where: {
                date: {
                  [Op.gte]: new Date(formattedFirstDay),
                  [Op.lte]: new Date(formattedLastDay),
                },
              },
              include: [{ model: Models.Name, as: "Name" }],
            },
          ],
        })
      );
    }

    await Promise.all(userData)
      .then((result) => {
        promiseUserData = result;
      })
      .catch((error) => {
        console.log(error);
      });

    let filterredData = [];
    for (let i = 0; i < promiseUserData.length; i++) {
      promiseUserData[i].map((item) => {
        if (item.dataValues.Data.length !== 0) {
          filterredData.push(item);
        }
      });
    }

    for (let i = 0; i < filterredData.length; i++) {
      let filterData = [];

      filterredData[i].Data.map((item) => {
        return filterData.push({
          name: item.Name.name,
          category: item.Name.category,
          date: item.date,
          amount: item.amount,
        });
      });

      let nameTotal = {};
      let categoryTotal = {};

      filterData.map((item) => {
        if (Object.hasOwn(nameTotal, `Total ${item.name}`)) {
          nameTotal[`Total ${item.name}`] += item.amount;
        } else {
          nameTotal[`Total ${item.name}`] = item.amount;
        }

        if (Object.hasOwn(categoryTotal, `Total ${item.category}`)) {
          categoryTotal[`Total ${item.category}`] += item.amount;
        } else {
          categoryTotal[`Total ${item.category}`] = item.amount;
        }
      });

      const total = filterData.reduce((sum, item) => sum + item.amount, 0);

      let solvedFormulas = await calculateFormulas(filterredData[i].id);

      const workbook = xlsx.utils.book_new();

      let worksheet;

      if (solvedFormulas === false) {
        worksheet = xlsx.utils.aoa_to_sheet([
          ["name", "category", "date", "amount"],
          ...filterData.map((item) => [
            item.name,
            item.category,
            item.date,
            item.amount,
          ]),
          [],
          ["Grand Total", total],
          [],
          ...Object.entries(categoryTotal).map(([key, value]) => [key, value]),
          [],
          ...Object.entries(nameTotal).map(([key, value]) => [key, value]),
        ]);
      } else {
        worksheet = xlsx.utils.aoa_to_sheet([
          ["name", "category", "date", "amount"],
          ...filterData.map((item) => [
            item.name,
            item.category,
            item.date,
            item.amount,
          ]),
          [],
          ["Grand Total", total],
          [],
          ...Object.entries(categoryTotal).map(([key, value]) => [key, value]),
          [],
          ...Object.entries(nameTotal).map(([key, value]) => [key, value]),
          [],
          ["Formula", "CalculatedValue"],
          ...solvedFormulas,
        ]);
      }

      var wscols = [
        { wch: 15 },
        { wch: 15 },
        { wch: 10 },
        { wch: 10 },
        { wch: 10 },
      ];
      worksheet["!cols"] = wscols;
      xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");
      xlsx.writeFile(workbook, `${filterredData[i].username}MonthlyData.xlsx`);
    }
  } catch (error) {
    console.log(`${messages.somethingWentWrong} : ${error}`);
  }
});

//1 10 * * MON

cron.schedule("1 10 * * MON", async () => {
  try {
    function getLastWeekDates() {
      var today = new Date();
      console.log("today: ", today);
      var lastWeek = new Date(
        today.getFullYear(),
        today.getMonth(),
        today.getDate() - 7
      );

      var lastWeekMonth = lastWeek.getMonth() + 1;
      var lastWeekDay = lastWeek.getDate();
      var lastWeekYear = lastWeek.getFullYear();

      var lastWeekMonday =
        lastWeekYear + "-" + lastWeekMonth + "-" + lastWeekDay;

      let lastWeekSunday =
        lastWeekYear + "-" + lastWeekMonth + "-" + (lastWeekDay + 6);

      return { lastWeekMonday, lastWeekSunday };
    }

    let { lastWeekMonday, lastWeekSunday } = getLastWeekDates();

    console.log("last week monday: ", lastWeekMonday);
    console.log("last week sunday: ", lastWeekSunday);

    const users = await Models.User.count({});

    let offset = users / 10;
    offset = Math.ceil(offset);

    let userData = [];
    let promiseUserData = [];
    for (let i = 0; i < offset; i++) {
      let limit = 10;
      userData.push(
        Models.User.findAll({
          limit,
          offset: i * limit,
          include: [
            {
              model: Models.Data,
              as: "Data",
              where: {
                date: {
                  [Op.gte]: new Date(lastWeekMonday),
                  [Op.lte]: new Date(lastWeekSunday),
                },
              },
              include: [{ model: Models.Name, as: "Name" }],
            },
          ],
        })
      );
    }

    await Promise.all(userData)
      .then((result) => {
        // console.log(result);
        promiseUserData = result;
      })
      .catch((error) => {
        console.log(error);
      });

    let filterredData = [];
    for (let i = 0; i < promiseUserData.length; i++) {
      promiseUserData[i].map((item) => {
        if (item.dataValues.Data.length !== 0) {
          filterredData.push(item);
        }
      });
      // console.log("Promise user data: ", promiseUserData[i]);
    }

    for (let i = 0; i < filterredData.length; i++) {
      let filterData = [];

      console.log("Consol;e", filterredData[i].id);

      filterredData[i].Data.map((item) => {
        return filterData.push({
          name: item.Name.name,
          category: item.Name.category,
          date: item.date,
          amount: item.amount,
        });
      });

      let nameTotal = {};
      let categoryTotal = {};

      filterData.map((item) => {
        if (Object.hasOwn(nameTotal, `Total ${item.name}`)) {
          nameTotal[`Total ${item.name}`] += item.amount;
        } else {
          nameTotal[`Total ${item.name}`] = item.amount;
        }

        if (Object.hasOwn(categoryTotal, `Total ${item.category}`)) {
          categoryTotal[`Total ${item.category}`] += item.amount;
        } else {
          categoryTotal[`Total ${item.category}`] = item.amount;
        }
      });

      const total = filterData.reduce((sum, item) => sum + item.amount, 0);

      let solvedFormulas = await calculateFormulas(filterredData[i].id);

      const workbook = xlsx.utils.book_new();

      let worksheet;

      if (solvedFormulas === false) {
        worksheet = xlsx.utils.aoa_to_sheet([
          ["name", "category", "date", "amount"],
          ...filterData.map((item) => [
            item.name,
            item.category,
            item.date,
            item.amount,
          ]),
          [],
          ["Grand Total", total],
          [],
          ...Object.entries(categoryTotal).map(([key, value]) => [key, value]),
          [],
          ...Object.entries(nameTotal).map(([key, value]) => [key, value]),
        ]);
      } else {
        worksheet = xlsx.utils.aoa_to_sheet([
          ["name", "category", "date", "amount"],
          ...filterData.map((item) => [
            item.name,
            item.category,
            item.date,
            item.amount,
          ]),
          [],
          ["Grand Total", total],
          [],
          ...Object.entries(categoryTotal).map(([key, value]) => [key, value]),
          [],
          ...Object.entries(nameTotal).map(([key, value]) => [key, value]),
          [],
          ["Formula", "CalculatedValue"],
          ...solvedFormulas,
        ]);
      }

      var wscols = [
        { wch: 15 },
        { wch: 15 },
        { wch: 10 },
        { wch: 10 },
        { wch: 10 },
      ];
      worksheet["!cols"] = wscols;
      xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");
      xlsx.writeFile(workbook, `${filterredData[i].username}WeeklyData.xlsx`);
    }
  } catch (error) {
    console.log(`${messages.somethingWentWrong} : ${error}`);
  }
});

//1 10 1 * *

cron.schedule("*/3 * * * * *", async () => {
  try {
    function formatFirstDate(date, format) {
      const map = {
        mm: date.getMonth(),
        dd: date.getDate(),
        yy: date.getFullYear(),
      };

      return format.replace(/mm|dd|yy|yyy/gi, (matched) => map[matched]);
    }

    function formatLastDate(date, format) {
      const map = {
        mm: date.getMonth(),
        dd: date.getDate(),
        yy: date.getFullYear(),
      };

      map.mm += 1;

      return format.replace(/mm|dd|yy|yyy/gi, (matched) => map[matched]);
    }

    function getLastMonthDate() {
      let date = new Date("2024-03-01T00:00:00.000Z"),
        y = date.getFullYear(),
        m = date.getMonth();

      let firstDay = new Date(y, m, 1);
      let lastDay = new Date(y, m, 0);

      return { firstDay, lastDay };
    }

    function getLastToLastMonthDate() {
      let date = new Date("2024-03-01T00:00:00.000Z"),
        y = date.getFullYear(),
        m = date.getMonth();

      let lastMonthFirstDay = new Date(y, m - 1, 1);
      let lastMonthLastDay = new Date(y, m - 1, 0);

      return { lastMonthFirstDay, lastMonthLastDay };
    }

    const { firstDay, lastDay } = getLastMonthDate();
    const lastMonthName = lastDay.toLocaleString("default", { month: "long" });

    const { lastMonthFirstDay, lastMonthLastDay } = getLastToLastMonthDate();

    const lastToLastMonthName = lastMonthLastDay.toLocaleString("default", {
      month: "long",
    });

    let formattedFirstDay = formatFirstDate(firstDay, "yy-mm-dd");
    let formattedLastDay = formatLastDate(lastDay, "yy-mm-dd");
    // console.log("last month: ", formattedFirstDay, formattedLastDay);

    let formattedLastMonthFirstDay = formatFirstDate(
      lastMonthFirstDay,
      "yy-mm-dd"
    );
    let formattedLastMonthLastDay = formatLastDate(
      lastMonthLastDay,
      "yy-mm-dd"
    );
    // console.log(
    //   "last to last month: ",
    //   formattedLastMonthFirstDay,
    //   formattedLastMonthLastDay
    // );

    const users = await Models.User.count({});

    let offset = users / 10;
    offset = Math.ceil(offset);

    formattedLastMonthFirstDay = new Date(formattedLastMonthFirstDay);

    formattedLastMonthFirstDay = new Date(
      formattedLastMonthFirstDay.getTime() +
        Math.abs(formattedLastMonthFirstDay.getTimezoneOffset() * 60000)
    );

    let userData = [];
    let promiseUserData = [];

    // console.log(new Date(formattedLastMonthFirstDay));
    // console.log(new Date(formattedLastDay));
    for (let i = 0; i < offset; i++) {
      let limit = 10;
      userData.push(
        Models.User.findAll({
          limit,
          offset: i * limit,
          include: [
            {
              model: Models.Data,
              as: "Data",
              where: {
                date: {
                  [Op.gte]: new Date(formattedLastMonthFirstDay),
                  [Op.lte]: new Date(formattedLastDay),
                },
              },
              include: [{ model: Models.Name, as: "Name" }],
            },
          ],
        })
      );
    }

    await Promise.all(userData)
      .then((result) => {
        // console.log(JSON.stringify(result));
        promiseUserData = result;
      })
      .catch((error) => {
        console.log(error);
      });

    for (let i = 0; i < promiseUserData.length; i++) {
      if (promiseUserData[i].length !== 0) {
        await promiseUserData[i].forEach(async (item) => {
          // console.log("Item: ", item.Data);
          const lastMonthData = item.Data.filter((data) => {
            const date = new Date(data.date);

            let formattedDate = new Date(formattedFirstDay);
            formattedDate = new Date(
              formattedDate.getTime() +
                Math.abs(formattedDate.getTimezoneOffset() * 60000)
            );

            return (
              date <= new Date(formattedLastDay) &&
              date >= new Date(formattedDate)
            );
          });

          const lastToLastMonthData = item.Data.filter((data) => {
            const date = new Date(data.date);

            let formattedFirstDate = new Date(formattedLastMonthFirstDay);
            formattedFirstDate = new Date(
              formattedFirstDate.getTime() +
                Math.abs(formattedFirstDate.getTimezoneOffset() * 60000)
            );

            let formattedLastDate = new Date(formattedLastMonthLastDay);
            formattedLastDate = new Date(
              formattedLastDate.getTime() +
                Math.abs(formattedLastDate.getTimezoneOffset() * 60000)
            );

            // console.log(new Date(formattedLastDate));
            // console.log(new Date(formattedFirstDate));

            return (
              date <= new Date(formattedLastDate) &&
              date >= new Date(formattedFirstDate)
            );
          });

          const lastMonthCategoryTotal = {};
          const lastToLastMonthCategoryTotal = {};

          lastMonthData.forEach((data) => {
            if (data.Name && data.Name.category) {
              if (
                Object.hasOwn(
                  lastMonthCategoryTotal,
                  `Total ${data.Name.category}`
                )
              ) {
                lastMonthCategoryTotal[`Total ${data.Name.category}`] +=
                  data.amount;
              } else {
                lastMonthCategoryTotal[`Total ${data.Name.category}`] =
                  data.amount;
              }
            }
          });

          lastToLastMonthData.forEach((data) => {
            if (data.Name && data.Name.category) {
              if (
                Object.hasOwn(
                  lastToLastMonthCategoryTotal,
                  `Total ${data.Name.category}`
                )
              ) {
                lastToLastMonthCategoryTotal[`Total ${data.Name.category}`] +=
                  data.amount;
              } else {
                lastToLastMonthCategoryTotal[`Total ${data.Name.category}`] =
                  data.amount;
              }
            }
          });

          let solvedFormulas = await calculateFormulas(item.dataValues.id);

          // console.log("Solved formulas: ", solvedFormulas);

          const workbook = xlsx.utils.book_new();

          let worksheet;

          if (solvedFormulas === false || solvedFormulas === undefined) {
            worksheet = xlsx.utils.aoa_to_sheet([
              [
                "category",
                lastToLastMonthName,
                lastMonthName,
                "Absolute Difference",
                "Percentage Difference",
              ],
              ...Object.keys(lastMonthCategoryTotal).map((category) => [
                category,
                lastToLastMonthCategoryTotal[category],
                lastMonthCategoryTotal[category],
                lastMonthCategoryTotal[category] -
                  lastToLastMonthCategoryTotal[category],
                `${(
                  ((lastMonthCategoryTotal[category] -
                    lastToLastMonthCategoryTotal[category]) /
                    lastToLastMonthCategoryTotal[category]) *
                  100
                ).toFixed(2)}%`,
              ]),
            ]);
          } else {
            worksheet = xlsx.utils.aoa_to_sheet([
              [
                "category",
                lastToLastMonthName,
                lastMonthName,
                "Absolute Difference",
                "Percentage Difference",
              ],
              ...Object.keys(lastMonthCategoryTotal).map((category) => [
                category,
                lastToLastMonthCategoryTotal[category],
                lastMonthCategoryTotal[category],
                lastMonthCategoryTotal[category] -
                  lastToLastMonthCategoryTotal[category],
                `${(
                  ((lastMonthCategoryTotal[category] -
                    lastToLastMonthCategoryTotal[category]) /
                    lastToLastMonthCategoryTotal[category]) *
                  100
                ).toFixed(2)}%`,
              ]),
              [],
              ["Formula", "CalculatedValue"],
              ...solvedFormulas,
            ]);
          }

          var wscols = [
            { wch: 15 },
            { wch: 15 },
            { wch: 10 },
            { wch: 16 },
            { wch: 18 },
          ];

          worksheet["!cols"] = wscols;

          xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");

          xlsx.writeFile(
            workbook,
            `${item.dataValues.username}TwoMonthMOMData.xlsx`
          );
        });
      }
    }
  } catch (error) {
    console.log(`${messages.somethingWentWrong} : ${error}`);
  }
});

// 1 10 1 JAN *

cron.schedule("1 10 1 JAN *", async () => {
  try {
    function getLastYearDate() {
      let date = new Date(),
        y = date.getFullYear();

      let fDay = new Date(y - 1, 0, 1);
      let lDay = new Date(y, 0, 0);

      let lastYearFirstDay = new Date(
        fDay.getTime() + Math.abs(fDay.getTimezoneOffset() * 60000)
      );
      let lastYearLastDay = new Date(
        lDay.getTime() + Math.abs(lDay.getTimezoneOffset() * 60000)
      );

      let lastYear = y - 1;

      return { lastYearFirstDay, lastYearLastDay, lastYear };
    }

    function getLastToLastMonthDate() {
      let date = new Date(),
        y = date.getFullYear();

      let fDay = new Date(y - 2, 0, 1);
      let lDay = new Date(y - 1, 0, 0);

      let lastToLastYearFirstDay = new Date(
        fDay.getTime() + Math.abs(fDay.getTimezoneOffset() * 60000)
      );
      let lastToLastYearLastDay = new Date(
        lDay.getTime() + Math.abs(lDay.getTimezoneOffset() * 60000)
      );

      let lastToLastYear = y - 2;

      return { lastToLastYearFirstDay, lastToLastYearLastDay, lastToLastYear };
    }

    const { lastYearFirstDay, lastYearLastDay, lastYear } = getLastYearDate();

    const { lastToLastYearFirstDay, lastToLastYearLastDay, lastToLastYear } =
      getLastToLastMonthDate();

    // console.log("yearL: ", lastToLastYear);

    const users = await Models.User.count({});

    let offset = users / 10;
    offset = Math.ceil(offset);

    let userData = [];
    let promiseUserData = [];
    for (let i = 0; i < offset; i++) {
      let limit = 10;
      userData.push(
        Models.User.findAll({
          limit,
          offset: i * limit,
          include: [
            {
              model: Models.Data,
              as: "Data",
              where: {
                date: {
                  [Op.gte]: new Date(lastToLastYearFirstDay),
                  [Op.lte]: new Date(lastYearLastDay),
                },
              },
              include: [{ model: Models.Name, as: "Name" }],
            },
          ],
        })
      );
    }

    await Promise.all(userData)
      .then((result) => {
        // console.log(JSON.stringify(result));
        promiseUserData = result;
      })
      .catch((error) => {
        console.log(error);
      });

    for (let i = 0; i < promiseUserData.length; i++) {
      if (promiseUserData[i].length !== 0) {
        await promiseUserData[i].forEach(async (item) => {
          const lastYearData = item.Data.filter((data) => {
            const date = new Date(data.date);

            return (
              date <= new Date(lastYearLastDay) &&
              date >= new Date(lastYearFirstDay)
            );
          });

          const lastToLastYearData = item.Data.filter((data) => {
            const date = new Date(data.date);

            return (
              date <= new Date(lastToLastYearLastDay) &&
              date >= new Date(lastToLastYearFirstDay)
            );
          });

          const lastYearCategoryTotal = {};
          const lastToLastYearCategoryTotal = {};

          lastYearData.forEach((data) => {
            if (data.Name && data.Name.category) {
              if (
                Object.hasOwn(
                  lastYearCategoryTotal,
                  `Total ${data.Name.category}`
                )
              ) {
                lastYearCategoryTotal[`Total ${data.Name.category}`] +=
                  data.amount;
              } else {
                lastYearCategoryTotal[`Total ${data.Name.category}`] =
                  data.amount;
              }
            }
          });

          lastToLastYearData.forEach((data) => {
            if (data.Name && data.Name.category) {
              if (
                Object.hasOwn(
                  lastToLastYearCategoryTotal,
                  `Total ${data.Name.category}`
                )
              ) {
                lastToLastYearCategoryTotal[`Total ${data.Name.category}`] +=
                  data.amount;
              } else {
                lastToLastYearCategoryTotal[`Total ${data.Name.category}`] =
                  data.amount;
              }
            }
          });

          let solvedFormulas = await calculateFormulas(item.dataValues.id);

          const workbook = xlsx.utils.book_new();
          let worksheet;

          if (solvedFormulas === false) {
            // console.log("solved :", solvedFormulas);
            worksheet = xlsx.utils.aoa_to_sheet([
              [
                "category",
                lastToLastYear,
                lastYear,
                "Absolute Difference",
                "Percentage Difference",
              ],
              ...Object.keys(lastYearCategoryTotal).map((category) => [
                category,
                lastToLastYearCategoryTotal[category],
                lastYearCategoryTotal[category],
                lastYearCategoryTotal[category] -
                  lastToLastYearCategoryTotal[category],
                `${(
                  ((lastYearCategoryTotal[category] -
                    lastToLastYearCategoryTotal[category]) /
                    lastToLastYearCategoryTotal[category]) *
                  100
                ).toFixed(2)}%`,
              ]),
            ]);
          } else {
            // console.log("solved :", solvedFormulas);
            worksheet = xlsx.utils.aoa_to_sheet([
              [
                "category",
                lastToLastYear,
                lastYear,
                "Absolute Difference",
                "Percentage Difference",
              ],
              ...Object.keys(lastYearCategoryTotal).map((category) => [
                category,
                lastToLastYearCategoryTotal[category],
                lastYearCategoryTotal[category],
                lastYearCategoryTotal[category] -
                  lastToLastYearCategoryTotal[category],
                `${(
                  ((lastYearCategoryTotal[category] -
                    lastToLastYearCategoryTotal[category]) /
                    lastToLastYearCategoryTotal[category]) *
                  100
                ).toFixed(2)}%`,
              ]),
              [],
              ["Formula", "CalculatedValue"],
              ...solvedFormulas,
            ]);
          }

          let wscols = [
            { wch: 15 },
            { wch: 15 },
            { wch: 10 },
            { wch: 16 },
            { wch: 18 },
          ];

          worksheet["!cols"] = wscols;

          xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");

          xlsx.writeFile(
            workbook,
            `${item.dataValues.username}TwoYearMOMData.xlsx`
          );
        });
      }
    }
  } catch (error) {
    console.log(`${messages.somethingWentWrong} : ${error}`);
  }
});
