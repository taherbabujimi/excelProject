const Models = require("../models/index");
const { messages } = require("../services/messages");
const xlsx = require("xlsx");
const { Op } = require("sequelize");
const cron = require("node-cron");

// 1 10 1 * *

cron.schedule("1 10 1 * *", async () => {
  try {
    function formatFirstDate(date, format) {
      const map = {
        mm: date.getMonth(),
        dd: date.getDate(),
        yy: date.getFullYear(),
      };

      console.log("mm", map.mm);

      return format.replace(/mm|dd|yy|yyy/gi, (matched) => map[matched]);
    }

    function formatLastDate(date, format) {
      const map = {
        mm: date.getMonth(),
        dd: date.getDate(),
        yy: date.getFullYear(),
      };

      map.mm += 1;

      console.log("mm", map.mm);

      return format.replace(/mm|dd|yy|yyy/gi, (matched) => map[matched]);
    }

    function GFG_Fun() {
      let date = new Date(),
        y = date.getFullYear(),
        m = date.getMonth();

      console.log(date);

      let firstDay = new Date(y, m, 1);
      let lastDay = new Date(y, m, 0);

      return { firstDay, lastDay };
    }

    const { firstDay, lastDay } = GFG_Fun();

    console.log("normal first day: ", firstDay);
    console.log("normal lst date: ", lastDay);

    let formattedFirstDay = formatFirstDate(firstDay, "yy-mm-dd");
    let formattedLastDay = formatLastDate(lastDay, "yy-mm-dd");

    console.log("First day=", formattedFirstDay);
    console.log("Last day = ", formattedLastDay);

    const data = await Models.Data.findAll({
      where: {
        date: {
          [Op.gte]: new Date(formattedFirstDay),
          [Op.lte]: new Date(formattedLastDay),
        },
      },
    });

    let dataName = [];
    let dataNameResult = [];

    for (let i = 0; i < data.length; i++) {
      dataName.push(
        Models.Name.findOne({
          where: { id: data[i].dataValues.name },
        })
      );
    }

    await Promise.all(dataName)
      .then((result) => {
        // console.log(result);
        dataNameResult = result.map((item) => {
          return {
            name: item.dataValues.name,
            id: item.dataValues.id,
            category: item.dataValues.category,
          };
        });
      })
      .catch((error) => {
        console.log("Error: ", error);
      });

    let filteredData = [];

    let nameTotal = {};

    let total = 0;
    let totalA = 0;
    let totalB = 0;
    let totalC = 0;
    let totalD = 0;
    let totalE = 0;
    for (let i = 0; i < data.length; i++) {
      console.log(data[i].dataValues);
      console.log(dataNameResult[i]);
      filteredData.push({
        name: dataNameResult[i].name,
        category: dataNameResult[i].category,
        date: data[i].dataValues.date,
        amount: data[i].dataValues.amount,
      });

      if (dataNameResult[i].category === "a") {
        totalA += data[i].dataValues.amount;
      }

      if (dataNameResult[i].category === "b") {
        totalB += data[i].dataValues.amount;
      }

      if (dataNameResult[i].category === "c") {
        totalC += data[i].dataValues.amount;
      }

      if (dataNameResult[i].category === "d") {
        totalD += data[i].dataValues.amount;
      }

      if (dataNameResult[i].category === "e") {
        totalE += data[i].dataValues.amount;
      }

      total += data[i].dataValues.amount;

      let Name = dataNameResult[i].name;

      if (Object.hasOwn(nameTotal, `Total ${Name}`)) {
        nameTotal[`Total ${Name}`] += data[i].dataValues.amount;
      } else {
        nameTotal[`Total ${Name}`] = data[i].dataValues.amount;
      }
    }

    console.log(filteredData);

    const workbook = xlsx.utils.book_new();

    const worksheet = xlsx.utils.aoa_to_sheet([
      ["name", "category", "date", "amount"],
      ...filteredData.map((item) => [
        item.name,
        item.category,
        item.date,
        item.amount,
      ]),
      [],
      ["Grand Total", total],
      [],
      ["Total A", totalA],
      ["Total B", totalB],
      ["Total C", totalC],
      ["Total D", totalD],
      ["Total E", totalE],
      [],
      ...Object.entries(nameTotal).map(([key, value]) => [key, value]),
    ]);

    var wscols = [
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
    ];

    worksheet["!cols"] = wscols;

    xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");

    xlsx.writeFile(
      workbook,
      `${formattedFirstDay}To${formattedLastDay}Data.xlsx`
    );
  } catch (error) {
    console.log(`${messages.somethingWentWrong} : ${error}`);
  }
});

//1 10 * * MON

cron.schedule("1 10 * * MON", async () => {
  try {
    function getLastWeekDates() {
      var today = new Date("2024-11-11T06:39:00.605Z");
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

    const data = await Models.Data.findAll({
      where: {
        date: {
          [Op.gte]: new Date(lastWeekMonday),
          [Op.lte]: new Date(lastWeekSunday),
        },
      },
    });

    let dataName = [];
    let dataNameResult = [];

    for (let i = 0; i < data.length; i++) {
      dataName.push(
        Models.Name.findOne({
          where: { id: data[i].dataValues.name },
        })
      );
    }

    await Promise.all(dataName)
      .then((result) => {
        // console.log(result);
        dataNameResult = result.map((item) => {
          return {
            name: item.dataValues.name,
            id: item.dataValues.id,
            category: item.dataValues.category,
          };
        });
      })
      .catch((error) => {
        console.log("Error: ", error);
      });

    let filteredData = [];

    let nameTotal = {};

    let total = 0;
    let totalA = 0;
    let totalB = 0;
    let totalC = 0;
    let totalD = 0;
    let totalE = 0;
    for (let i = 0; i < data.length; i++) {
      filteredData.push({
        name: dataNameResult[i].name,
        category: dataNameResult[i].category,
        date: data[i].dataValues.date,
        amount: data[i].dataValues.amount,
      });

      if (dataNameResult[i].category === "a") {
        totalA += data[i].dataValues.amount;
      }

      if (dataNameResult[i].category === "b") {
        totalB += data[i].dataValues.amount;
      }

      if (dataNameResult[i].category === "c") {
        totalC += data[i].dataValues.amount;
      }

      if (dataNameResult[i].category === "d") {
        totalD += data[i].dataValues.amount;
      }

      if (dataNameResult[i].category === "e") {
        totalE += data[i].dataValues.amount;
      }

      total += data[i].dataValues.amount;

      let Name = dataNameResult[i].name;

      if (Object.hasOwn(nameTotal, `Total ${Name}`)) {
        nameTotal[`Total ${Name}`] += data[i].dataValues.amount;
      } else {
        nameTotal[`Total ${Name}`] = data[i].dataValues.amount;
      }
    }

    console.log(nameTotal);

    console.log("total", total);

    console.log(filteredData);

    const workbook = xlsx.utils.book_new();

    const worksheet = xlsx.utils.aoa_to_sheet([
      ["name", "category", "date", "amount"],
      ...filteredData.map((item) => [
        item.name,
        item.category,
        item.date,
        item.amount,
      ]),
      [],
      ["Grand Total", total],
      [],
      ["Total A", totalA],
      ["Total B", totalB],
      ["Total C", totalC],
      ["Total D", totalD],
      ["Total E", totalE],
      [],
      ...Object.entries(nameTotal).map(([key, value]) => [key, value]),
    ]);

    var wscols = [
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
    ];

    worksheet["!cols"] = wscols;

    xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");

    xlsx.writeFile(workbook, `${lastWeekMonday}To${lastWeekSunday}Data.xlsx`);
  } catch (error) {
    console.log(`${messages.somethingWentWrong} : ${error}`);
  }
});

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

    function getLastMonthDate() {
      let date = new Date(),
        y = date.getFullYear(),
        m = date.getMonth();

      let firstDay = new Date(y, m, 1);
      let lastDay = new Date(y, m, 0);

      return { firstDay, lastDay };
    }

    function getLastToLastMonthDate() {
      let date = new Date(),
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

    let formattedLastMonthFirstDay = formatFirstDate(
      lastMonthFirstDay,
      "yy-mm-dd"
    );
    let formattedLastMonthLastDay = formatLastDate(
      lastMonthLastDay,
      "yy-mm-dd"
    );

    const data = await Models.Data.findAll({
      where: {
        date: {
          [Op.gte]: new Date(formattedFirstDay),
          [Op.lte]: new Date(formattedLastDay),
        },
      },
    });

    const lastMonthData = await Models.Data.findAll({
      where: {
        date: {
          [Op.gte]: new Date(formattedLastMonthFirstDay),
          [Op.lte]: new Date(formattedLastMonthLastDay),
        },
      },
    });

    let dataName = [];
    let dataNameResult = [];

    for (let i = 0; i < data.length; i++) {
      dataName.push(
        Models.Name.findOne({
          where: { id: data[i].dataValues.name },
        })
      );
    }

    await Promise.all(dataName)
      .then((result) => {
        // console.log(result);
        dataNameResult = result.map((item) => {
          return {
            name: item.dataValues.name,
            id: item.dataValues.id,
            category: item.dataValues.category,
          };
        });
      })
      .catch((error) => {
        console.log("Error: ", error);
      });

    // console.log(dataNameResult);

    let lastMonthTotalA = 0;
    let lastMonthTotalB = 0;
    let lastMonthTotalC = 0;
    let lastMonthTotalD = 0;
    let lastMonthTotalE = 0;

    for (let i = 0; i < data.length; i++) {
      if (dataNameResult[i].category === "a") {
        lastMonthTotalA += data[i].dataValues.amount;
      }

      if (dataNameResult[i].category === "b") {
        lastMonthTotalB += data[i].dataValues.amount;
      }

      if (dataNameResult[i].category === "c") {
        lastMonthTotalC += data[i].dataValues.amount;
      }

      if (dataNameResult[i].category === "d") {
        lastMonthTotalD += data[i].dataValues.amount;
      }

      if (dataNameResult[i].category === "e") {
        lastMonthTotalE += data[i].dataValues.amount;
      }
    }

    let lastMonthDataName = [];
    let lastMonthDataNameResult = [];

    for (let i = 0; i < lastMonthData.length; i++) {
      lastMonthDataName.push(
        Models.Name.findOne({
          where: { id: data[i].dataValues.name },
        })
      );
    }

    await Promise.all(lastMonthDataName)
      .then((result) => {
        // console.log(result);
        lastMonthDataNameResult = result.map((item) => {
          return {
            name: item.dataValues.name,
            id: item.dataValues.id,
            category: item.dataValues.category,
          };
        });
      })
      .catch((error) => {
        console.log("Error: ", error);
      });

    let lastToLastMonthTotalA = 0;
    let lastToLastMonthTotalB = 0;
    let lastToLastMonthTotalC = 0;
    let lastToLastMonthTotalD = 0;
    let lastToLastMonthTotalE = 0;
    for (let i = 0; i < lastMonthData.length; i++) {
      if (lastMonthDataNameResult[i].category === "a") {
        lastToLastMonthTotalA += lastMonthData[i].dataValues.amount;
      }

      if (lastMonthDataNameResult[i].category === "b") {
        lastToLastMonthTotalB += lastMonthData[i].dataValues.amount;
      }

      if (lastMonthDataNameResult[i].category === "c") {
        lastToLastMonthTotalC += lastMonthData[i].dataValues.amount;
      }

      if (lastMonthDataNameResult[i].category === "d") {
        lastToLastMonthTotalD += lastMonthData[i].dataValues.amount;
      }

      if (lastMonthDataNameResult[i].category === "e") {
        lastToLastMonthTotalE += lastMonthData[i].dataValues.amount;
      }
    }

    const workbook = xlsx.utils.book_new();

    const worksheet = xlsx.utils.aoa_to_sheet([
      [
        "category",
        lastToLastMonthName,
        lastMonthName,
        "Absolute Difference",
        "Percentage Difference",
      ],
      [
        `Total A`,
        lastToLastMonthTotalA,
        lastMonthTotalA,
        lastMonthTotalA - lastToLastMonthTotalA,
        `${(
          (100 * (lastMonthTotalA - lastToLastMonthTotalA)) /
          ((lastMonthTotalA + lastToLastMonthTotalA) / 2)
        ).toFixed(2)}%`,
      ],
      [
        "Total B",
        lastToLastMonthTotalB,
        lastMonthTotalB,
        lastMonthTotalB - lastToLastMonthTotalB,
        `${(
          (100 * (lastMonthTotalB - lastToLastMonthTotalB)) /
          ((lastMonthTotalB + lastToLastMonthTotalB) / 2)
        ).toFixed(2)}%`,
      ],
      [
        "Total C",
        lastToLastMonthTotalC,
        lastMonthTotalC,
        lastMonthTotalC - lastToLastMonthTotalC,
        `${(
          (100 * (lastMonthTotalC - lastToLastMonthTotalC)) /
          ((lastMonthTotalC + lastToLastMonthTotalC) / 2)
        ).toFixed(2)}%`,
      ],
      [
        "Total D",
        lastToLastMonthTotalD,
        lastMonthTotalD,
        lastMonthTotalD - lastToLastMonthTotalD,
        `${(
          (100 * (lastMonthTotalD - lastToLastMonthTotalD)) /
          ((lastMonthTotalD + lastToLastMonthTotalD) / 2)
        ).toFixed(2)}%`,
      ],
      [
        "Total E",
        lastToLastMonthTotalE,
        lastMonthTotalE,
        lastMonthTotalE - lastToLastMonthTotalE,
        `${(
          (100 * (lastMonthTotalE - lastToLastMonthTotalE)) /
          ((lastMonthTotalE + lastToLastMonthTotalE) / 2)
        ).toFixed(2)}%`,
      ],
    ]);

    var wscols = [
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 16 },
      { wch: 18 },
    ];

    worksheet["!cols"] = wscols;

    xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");

    xlsx.writeFile(workbook, `twoMonthMOMData.xlsx`);
  } catch (error) {
    console.log(`${messages.somethingWentWrong} : ${error}`);
  }
});

//1 10 1 JAN *
cron.schedule("*/30 * * * * *", async () => {
  try {
    function formatFirstDate(date, format) {
      const map = {
        mm: date.getMonth() + 1,
        dd: date.getDate(),
        yy: date.getFullYear(),
      };

      console.log(format.replace(/mm|dd|yy|yyy/gi, (matched) => map[matched]));

      return format.replace(/mm|dd|yy|yyy/gi, (matched) => map[matched]);
    }

    function formatLastDate(date, format) {
      const map = {
        mm: date.getMonth(),
        dd: date.getDate(),
        yy: date.getFullYear(),
      };

      console.log("mm: ", map.mm);
      map.mm += 1;

      return format.replace(/mm|dd|yy|yyy/gi, (matched) => map[matched]);
    }

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

    console.log("last year: ", lastYear);
    console.log("last year first day: ", lastYearFirstDay);
    console.log("last year last day: ", lastYearLastDay);

    const { lastToLastYearFirstDay, lastToLastYearLastDay, lastToLastYear } =
      getLastToLastMonthDate();

    console.log("yearL: ", lastToLastYear);

    console.log("last to last :", lastToLastYearFirstDay);
    console.log("last to last :", lastToLastYearLastDay);

    let formattedLastYearFirstDay = formatFirstDate(
      lastYearFirstDay,
      "yy-mm-dd"
    );
    let formattedLastYearLastDay = formatLastDate(lastYearLastDay, "yy-mm-dd");

    console.log("formatted: ", formattedLastYearFirstDay);
    console.log("formatted: ", formattedLastYearLastDay);

    let formattedLastToLastYearFirstDay = formatFirstDate(
      lastToLastYearFirstDay,
      "yy-mm-dd"
    );
    let formattedLastToLastYearLastDay = formatLastDate(
      lastToLastYearLastDay,
      "yy-mm-dd"
    );

    console.log("formatted: ", formattedLastToLastYearFirstDay);
    console.log("formatted: ", formattedLastToLastYearLastDay);

    const data = await Models.Data.findAll({
      where: {
        date: {
          [Op.gt]: new Date(formattedLastYearFirstDay),
          [Op.lte]: new Date(formattedLastYearLastDay),
        },
      },
    });

    const lastMonthData = await Models.Data.findAll({
      where: {
        date: {
          [Op.gt]: new Date(formattedLastToLastYearFirstDay),
          [Op.lte]: new Date(formattedLastToLastYearLastDay),
        },
      },
    });

    let dataName = [];
    let dataNameResult = [];

    for (let i = 0; i < data.length; i++) {
      dataName.push(
        Models.Name.findOne({
          where: { id: data[i].dataValues.name },
        })
      );
    }

    await Promise.all(dataName)
      .then((result) => {
        // console.log(result);
        dataNameResult = result.map((item) => {
          return {
            name: item.dataValues.name,
            id: item.dataValues.id,
            category: item.dataValues.category,
          };
        });
      })
      .catch((error) => {
        console.log("Error: ", error);
      });

    // console.log(dataNameResult);

    let lastMonthTotalA = 0;
    let lastMonthTotalB = 0;
    let lastMonthTotalC = 0;
    let lastMonthTotalD = 0;
    let lastMonthTotalE = 0;

    for (let i = 0; i < data.length; i++) {
      if (dataNameResult[i].category === "a") {
        lastMonthTotalA += data[i].dataValues.amount;
      }

      if (dataNameResult[i].category === "b") {
        lastMonthTotalB += data[i].dataValues.amount;
      }

      if (dataNameResult[i].category === "c") {
        lastMonthTotalC += data[i].dataValues.amount;
      }

      if (dataNameResult[i].category === "d") {
        lastMonthTotalD += data[i].dataValues.amount;
      }

      if (dataNameResult[i].category === "e") {
        lastMonthTotalE += data[i].dataValues.amount;
      }
    }

    let lastMonthDataName = [];
    let lastMonthDataNameResult = [];

    for (let i = 0; i < lastMonthData.length; i++) {
      lastMonthDataName.push(
        Models.Name.findOne({
          where: { id: data[i].dataValues.name },
        })
      );
    }

    await Promise.all(lastMonthDataName)
      .then((result) => {
        // console.log(result);
        lastMonthDataNameResult = result.map((item) => {
          return {
            name: item.dataValues.name,
            id: item.dataValues.id,
            category: item.dataValues.category,
          };
        });
      })
      .catch((error) => {
        console.log("Error: ", error);
      });

    let lastToLastMonthTotalA = 0;
    let lastToLastMonthTotalB = 0;
    let lastToLastMonthTotalC = 0;
    let lastToLastMonthTotalD = 0;
    let lastToLastMonthTotalE = 0;
    for (let i = 0; i < lastMonthData.length; i++) {
      if (lastMonthDataNameResult[i].category === "a") {
        lastToLastMonthTotalA += lastMonthData[i].dataValues.amount;
      }

      if (lastMonthDataNameResult[i].category === "b") {
        lastToLastMonthTotalB += lastMonthData[i].dataValues.amount;
      }

      if (lastMonthDataNameResult[i].category === "c") {
        lastToLastMonthTotalC += lastMonthData[i].dataValues.amount;
      }

      if (lastMonthDataNameResult[i].category === "d") {
        lastToLastMonthTotalD += lastMonthData[i].dataValues.amount;
      }

      if (lastMonthDataNameResult[i].category === "e") {
        lastToLastMonthTotalE += lastMonthData[i].dataValues.amount;
      }
    }

    const workbook = xlsx.utils.book_new();

    const worksheet = xlsx.utils.aoa_to_sheet([
      [
        "category",
        lastToLastYear,
        lastYear,
        "Absolute Difference",
        "Percentage Difference",
      ],
      [
        `Total A`,
        lastToLastMonthTotalA,
        lastMonthTotalA,
        lastMonthTotalA - lastToLastMonthTotalA,
        `${(
          (100 * (lastMonthTotalA - lastToLastMonthTotalA)) /
          ((lastMonthTotalA + lastToLastMonthTotalA) / 2)
        ).toFixed(2)}%`,
      ],
      [
        "Total B",
        lastToLastMonthTotalB,
        lastMonthTotalB,
        lastMonthTotalB - lastToLastMonthTotalB,
        `${(
          (100 * (lastMonthTotalB - lastToLastMonthTotalB)) /
          ((lastMonthTotalB + lastToLastMonthTotalB) / 2)
        ).toFixed(2)}%`,
      ],
      [
        "Total C",
        lastToLastMonthTotalC,
        lastMonthTotalC,
        lastMonthTotalC - lastToLastMonthTotalC,
        `${(
          (100 * (lastMonthTotalC - lastToLastMonthTotalC)) /
          ((lastMonthTotalC + lastToLastMonthTotalC) / 2)
        ).toFixed(2)}%`,
      ],
      [
        "Total D",
        lastToLastMonthTotalD,
        lastMonthTotalD,
        lastMonthTotalD - lastToLastMonthTotalD,
        `${(
          (100 * (lastMonthTotalD - lastToLastMonthTotalD)) /
          ((lastMonthTotalD + lastToLastMonthTotalD) / 2)
        ).toFixed(2)}%`,
      ],
      [
        "Total E",
        lastToLastMonthTotalE,
        lastMonthTotalE,
        lastMonthTotalE - lastToLastMonthTotalE,
        `${(
          (100 * (lastMonthTotalE - lastToLastMonthTotalE)) /
          ((lastMonthTotalE + lastToLastMonthTotalE) / 2)
        ).toFixed(2)}%`,
      ],
    ]);

    var wscols = [
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 16 },
      { wch: 18 },
    ];

    worksheet["!cols"] = wscols;

    xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");

    xlsx.writeFile(workbook, `twoMonthMOMData.xlsx`);
  } catch (error) {
    console.log(`${messages.somethingWentWrong} : ${error}`);
  }
});
