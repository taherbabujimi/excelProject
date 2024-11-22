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
      include: [{ model: Models.Name, as: "Name" }],
    });

    let filteredData = data.map((item) => {
      return {
        name: item.dataValues.Name.dataValues.name,
        category: item.dataValues.Name.dataValues.category,
        date: item.dataValues.date,
        amount: item.dataValues.amount,
      };
    });

    const categoryTotal = await Models.Data.findAll({
      where: {
        date: {
          [Op.gte]: new Date(formattedFirstDay),
          [Op.lte]: new Date(formattedLastDay),
        },
      },
      include: [
        {
          model: Models.Name,
          as: "Name",
          attributes: [],
        },
      ],
      attributes: [
        [
          Models.Sequelize.fn("SUM", Models.Sequelize.col("amount")),
          "totalAmount",
        ],
        [Models.Sequelize.col("Name.category"), "category"],
      ],
      group: ["Name.category"],
      raw: true,
    });

    console.log("category total: ", categoryTotal);

    const total = data.reduce((sum, item) => sum + item.amount, 0);

    let nameTotal = {};

    filteredData.map((item) => {
      if (Object.hasOwn(nameTotal, `Total ${item.name}`)) {
        nameTotal[`Total ${item.name}`] += item.amount;
      } else {
        nameTotal[`Total ${item.name}`] = item.amount;
      }
    });

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
      ["Total A", categoryTotal[0]?.totalAmount.toFixed(2)],
      ["Total B", categoryTotal[1]?.totalAmount.toFixed(2)],
      ["Total C", categoryTotal[2]?.totalAmount.toFixed(2)],
      ["Total D", categoryTotal[3]?.totalAmount.toFixed(2)],
      ["Total E", categoryTotal[4]?.totalAmount.toFixed(2)],
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
      include: [{ model: Models.Name, as: "Name" }],
    });

    let filteredData = data.map((item) => {
      return {
        name: item.dataValues.Name.dataValues.name,
        category: item.dataValues.Name.dataValues.category,
        date: item.dataValues.date,
        amount: item.dataValues.amount,
      };
    });

    const categoryTotal = await Models.Data.findAll({
      where: {
        date: {
          [Op.gte]: new Date(lastWeekMonday),
          [Op.lte]: new Date(lastWeekSunday),
        },
      },
      include: [
        {
          model: Models.Name,
          as: "Name",
          attributes: [],
        },
      ],
      attributes: [
        [
          Models.Sequelize.fn("SUM", Models.Sequelize.col("amount")),
          "totalAmount",
        ],
        [Models.Sequelize.col("Name.category"), "category"],
      ],
      group: ["Name.category"],
      raw: true,
    });

    console.log("category total: ", categoryTotal);

    const total = data.reduce((sum, item) => sum + item.amount, 0);

    let nameTotal = {};

    filteredData.map((item) => {
      if (Object.hasOwn(nameTotal, `Total ${item.name}`)) {
        nameTotal[`Total ${item.name}`] += item.amount;
      } else {
        nameTotal[`Total ${item.name}`] = item.amount;
      }
    });

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
      ["Total A", categoryTotal[0]?.totalAmount.toFixed(2)],
      ["Total B", categoryTotal[1]?.totalAmount.toFixed(2)],
      ["Total C", categoryTotal[2]?.totalAmount.toFixed(2)],
      ["Total D", categoryTotal[3]?.totalAmount.toFixed(2)],
      ["Total E", categoryTotal[4]?.totalAmount.toFixed(2)],
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
    console.log("last month: ",formattedFirstDay, formattedLastDay);

    let formattedLastMonthFirstDay = formatFirstDate(
      lastMonthFirstDay,
      "yy-mm-dd"
    );
    let formattedLastMonthLastDay = formatLastDate(
      lastMonthLastDay,
      "yy-mm-dd"
    );
    console.log("last to last month: ",formattedLastMonthFirstDay, formattedLastMonthLastDay);

    const lastMonthCategoryTotal = await Models.Data.findAll({
      where: {
        date: {
          [Op.gte]: new Date(formattedFirstDay),
          [Op.lte]: new Date(formattedLastDay),
        },
      },
      include: [
        {
          model: Models.Name,
          as: "Name",
          attributes: [],
        },
      ],
      attributes: [
        [
          Models.Sequelize.fn("SUM", Models.Sequelize.col("amount")),
          "totalAmount",
        ],
        [Models.Sequelize.col("Name.category"), "category"],
      ],
      group: ["Name.category"],
      raw: true,
    });

    console.log("october category total: ", lastMonthCategoryTotal);

    const lastToLastMonthCategoryTotal = await Models.Data.findAll({
      where: {
        date: {
          [Op.gte]: new Date(formattedLastMonthFirstDay),
          [Op.lte]: new Date(formattedLastMonthLastDay),
        },
      },
      include: [
        {
          model: Models.Name,
          as: "Name",
          attributes: [],
        },
      ],
      attributes: [
        [
          Models.Sequelize.fn("SUM", Models.Sequelize.col("amount")),
          "totalAmount",
        ],
        [Models.Sequelize.col("Name.category"), "category"],
      ],
      group: ["Name.category"],
      raw: true,
    });

    console.log("september: ", lastToLastMonthCategoryTotal);

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
        lastToLastMonthCategoryTotal[0].totalAmount.toFixed(2),
        lastMonthCategoryTotal[0].totalAmount.toFixed(2),
        lastMonthCategoryTotal[0].totalAmount.toFixed(2) -
          lastToLastMonthCategoryTotal[0].totalAmount.toFixed(2),
        `${(
          (100 *
            (lastMonthCategoryTotal[0].totalAmount -
              lastToLastMonthCategoryTotal[0].totalAmount)) /
          ((lastMonthCategoryTotal[0].totalAmount +
            lastToLastMonthCategoryTotal[0].totalAmount) /
            2)
        ).toFixed(2)}%`,
      ],
      [
        "Total B",
        lastToLastMonthCategoryTotal[1].totalAmount.toFixed(2),
        lastMonthCategoryTotal[1].totalAmount.toFixed(2),
        lastMonthCategoryTotal[1].totalAmount.toFixed(2) -
          lastToLastMonthCategoryTotal[1].totalAmount.toFixed(2),
        `${(
          (100 *
            (lastMonthCategoryTotal[1].totalAmount -
              lastToLastMonthCategoryTotal[1].totalAmount)) /
          ((lastMonthCategoryTotal[1].totalAmount +
            lastToLastMonthCategoryTotal[1].totalAmount) /
            2)
        ).toFixed(2)}%`,
      ],
      [
        "Total C",
        lastToLastMonthCategoryTotal[2].totalAmount.toFixed(2),
        lastMonthCategoryTotal[2].totalAmount.toFixed(2),
        lastMonthCategoryTotal[2].totalAmount.toFixed(2) -
          lastToLastMonthCategoryTotal[2].totalAmount.toFixed(2),
        `${(
          (100 *
            (lastMonthCategoryTotal[2].totalAmount -
              lastToLastMonthCategoryTotal[2].totalAmount)) /
          ((lastMonthCategoryTotal[2].totalAmount +
            lastToLastMonthCategoryTotal[2].totalAmount) /
            2)
        ).toFixed(2)}%`,
      ],
      [
        "Total D",
        lastToLastMonthCategoryTotal[3].totalAmount.toFixed(2),
        lastMonthCategoryTotal[3].totalAmount.toFixed(2),
        lastMonthCategoryTotal[3].totalAmount.toFixed(2) -
          lastToLastMonthCategoryTotal[3].totalAmount.toFixed(2),
        `${(
          (100 *
            (lastMonthCategoryTotal[3].totalAmount -
              lastToLastMonthCategoryTotal[3].totalAmount)) /
          ((lastMonthCategoryTotal[3].totalAmount +
            lastToLastMonthCategoryTotal[3].totalAmount) /
            2)
        ).toFixed(2)}%`,
      ],
      [
        "Total E",
        lastToLastMonthCategoryTotal[4].totalAmount.toFixed(2),
        lastMonthCategoryTotal[4].totalAmount.toFixed(2),
        lastMonthCategoryTotal[4].totalAmount.toFixed(2) -
          lastToLastMonthCategoryTotal[4].totalAmount.toFixed(2),
        `${(
          (100 *
            (lastMonthCategoryTotal[4].totalAmount -
              lastToLastMonthCategoryTotal[4].totalAmount)) /
          ((lastMonthCategoryTotal[4].totalAmount +
            lastToLastMonthCategoryTotal[4].totalAmount) /
            2)
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

// 1 10 1 JAN *
cron.schedule("1 10 1 JAN *", async () => {
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

    const { lastToLastYearFirstDay, lastToLastYearLastDay, lastToLastYear } =
      getLastToLastMonthDate();

    console.log("yearL: ", lastToLastYear);

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

    const lastYearCategoryTotal = await Models.Data.findAll({
      where: {
        date: {
          [Op.gte]: new Date(formattedLastYearFirstDay),
          [Op.lte]: new Date(formattedLastYearLastDay),
        },
      },
      include: [
        {
          model: Models.Name,
          as: "Name",
          attributes: [],
        },
      ],
      attributes: [
        [
          Models.Sequelize.fn("SUM", Models.Sequelize.col("amount")),
          "totalAmount",
        ],
        [Models.Sequelize.col("Name.category"), "category"],
      ],
      group: ["Name.category"],
      raw: true,
    });
    console.log(lastYearCategoryTotal);

    //------------------------------------------------------------------------

    const lastToLastYearCategoryTotal = await Models.Data.findAll({
      where: {
        date: {
          [Op.gte]: new Date(formattedLastToLastYearFirstDay),
          [Op.lte]: new Date(formattedLastToLastYearLastDay),
        },
      },
      include: [
        {
          model: Models.Name,
          as: "Name",
          attributes: [],
        },
      ],
      attributes: [
        [
          Models.Sequelize.fn("SUM", Models.Sequelize.col("amount")),
          "totalAmount",
        ],
        [Models.Sequelize.col("Name.category"), "category"],
      ],
      group: ["Name.category"],
      raw: true,
    });

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
        lastToLastYearCategoryTotal[0].totalAmount,
        lastYearCategoryTotal[0].totalAmount,
        lastYearCategoryTotal[0].totalAmount -
          lastToLastYearCategoryTotal[0].totalAmount,
        `${(
          (100 *
            (lastYearCategoryTotal[0].totalAmount -
              lastToLastYearCategoryTotal[0].totalAmount)) /
          ((lastYearCategoryTotal[0].totalAmount +
            lastToLastYearCategoryTotal[0].totalAmount) /
            2)
        ).toFixed(2)}%`,
      ],
      [
        "Total B",
        lastToLastYearCategoryTotal[1].totalAmount,
        lastYearCategoryTotal[1].totalAmount,
        lastYearCategoryTotal[1].totalAmount -
          lastToLastYearCategoryTotal[1].totalAmount,
        `${(
          (100 *
            (lastYearCategoryTotal[1].totalAmount -
              lastToLastYearCategoryTotal[1].totalAmount)) /
          ((lastYearCategoryTotal[1].totalAmount +
            lastToLastYearCategoryTotal[1].totalAmount) /
            2)
        ).toFixed(2)}%`,
      ],
      [
        "Total C",
        lastToLastYearCategoryTotal[2].totalAmount,
        lastYearCategoryTotal[2].totalAmount,
        lastYearCategoryTotal[2].totalAmount -
          lastToLastYearCategoryTotal[2].totalAmount,
        `${(
          (100 *
            (lastYearCategoryTotal[2].totalAmount -
              lastToLastYearCategoryTotal[2].totalAmount)) /
          ((lastYearCategoryTotal[2].totalAmount +
            lastToLastYearCategoryTotal[2].totalAmount) /
            2)
        ).toFixed(2)}%`,
      ],
      [
        "Total D",
        lastToLastYearCategoryTotal[3].totalAmount,
        lastYearCategoryTotal[3].totalAmount,
        lastYearCategoryTotal[3].totalAmount -
          lastToLastYearCategoryTotal[3].totalAmount,
        `${(
          (100 *
            (lastYearCategoryTotal[3].totalAmount -
              lastToLastYearCategoryTotal[3].totalAmount)) /
          ((lastYearCategoryTotal[3].totalAmount +
            lastToLastYearCategoryTotal[3].totalAmount) /
            2)
        ).toFixed(2)}%`,
      ],
      [
        "Total E",
        lastToLastYearCategoryTotal[4].totalAmount,
        lastYearCategoryTotal[4].totalAmount,
        lastYearCategoryTotal[4].totalAmount -
          lastToLastYearCategoryTotal[4].totalAmount,
        `${(
          (100 *
            (lastYearCategoryTotal[4].totalAmount -
              lastToLastYearCategoryTotal[4].totalAmount)) /
          ((lastYearCategoryTotal[4].totalAmount +
            lastToLastYearCategoryTotal[4].totalAmount) /
            2)
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

    xlsx.writeFile(workbook, `twoYearMOMData.xlsx`);
  } catch (error) {
    console.log(`${messages.somethingWentWrong} : ${error}`);
  }
});
