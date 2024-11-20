const Models = require("../models/index");
const { sequelize } = require("../models/index");
const {
  addDataSchema,
  getDataSchema,
} = require("../services/validation/dataValidation");
const {
  successResponseData,
  errorResponseData,
  errorResponseWithoutData,
  successResponseWithoutData,
} = require("../services/responses");
const { messages } = require("../services/messages");
const xlsx = require("xlsx");
const { Op } = require("sequelize");

function parseDate(dateStr) {
  let day, month, year;

  const dashFormat = /^\d{2}-\d{2}-\d{4}$/.test(dateStr);

  const slashFormat = /^\d{1,2}\/\d{1,2}\/\d{2}$/.test(dateStr);

  let dateParts = 0;

  if (dashFormat) {
    dateParts = dateStr.split("-");
  } else if (slashFormat) {
    dateParts = dateStr.split("/");
  } else {
    return null;
  }

  day = parseInt(dateParts[0], 10);
  month = parseInt(dateParts[1], 10) - 1;
  year = parseInt(dateParts[2], 10);

  if (year < 100) {
    year += 2000;
  }

  return new Date(year, month, day);
}

module.exports.addData = async (req, res) => {
  const transaction = await sequelize.transaction();
  let hasRolledBack = false;

  try {
    const workbook = xlsx.read(req.file.buffer, {
      type: "buffer",
    });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const excelData = xlsx.utils.sheet_to_json(worksheet, { raw: false });

    if (excelData.length === 0) {
      return errorResponseWithoutData(res, messages.excelFileEmpty, 400);
    }

    let namePromises = [];
    let existedNames = [];
    for (const row of excelData) {
      let { name } = row;
      name = name.toLowerCase();
      let arr = name.split(" ").filter((x) => x !== "");
      let finalName = "";

      for (let item of arr) {
        finalName += item + " ";
      }

      finalName = finalName.charAt(0).toUpperCase() + finalName.slice(1);
      finalName = finalName.trim();

      namePromises.push(
        Models.Name.findOne(
          {
            where: { [Op.and]: [{ name: finalName }, { userId: req.user.id }] },
          },
          { transaction }
        )
      );
    }
    await Promise.all(namePromises)
      .then((result) => {
        existedNames = result
          .filter((item) => item !== null)
          .map((item) => {
            return { name: item.dataValues.name, id: item.dataValues.id };
          });
      })
      .catch((error) => {
        console.log("error: ", error);
      });

    const createdData = [];
    const errorsInRow = [];
    const processedNames = new Set();
    let count = 0;

    let updatedData = [];
    let newData = [];
    let newNamePromises = [];
    const uniqueNames = new Set();

    for (const row of excelData) {
      count++;

      let keys = Object.keys(row);

      const allowedValues = ["name", "category", "date", "amount"];

      if (keys.some((value) => !allowedValues.includes(value))) {
        return errorResponseWithoutData(
          res,
          messages.provideAllowedColumns,
          400
        );
      }

      let { name, category, date, amount } = row;
      if ([name, category, date, amount].some((field) => field === undefined)) {
        errorsInRow.push([
          `Row:${count}: ${messages.provideFieldAndColumnProperly}`,
        ]);
        continue;
      }

      let formattedDate = parseDate(date);
      formattedDate = new Date(
        formattedDate.getTime() +
          Math.abs(formattedDate.getTimezoneOffset() * 60000)
      );

      name = name.toLowerCase();
      let arr = name.split(" ").filter((x) => x !== "");
      let finalName = "";

      for (let item of arr) {
        finalName += item + " ";
      }

      finalName = finalName.charAt(0).toUpperCase() + finalName.slice(1);
      finalName = finalName.trim();

      const body = { name: finalName, category, date: formattedDate, amount };

      const validationResponse = addDataSchema(body, res);
      if (validationResponse !== false) {
        errorsInRow.push([`Row:${count}: ${validationResponse}`]);
        continue;
      }

      console.log("Processed Names: ", processedNames);

      const duplicateEntry = processedNames.has(
        `${finalName}-${formattedDate}`
      );
      if (duplicateEntry) {
        errorsInRow.push([
          `Row:${count}: Duplicate entry for name: ${finalName} on date: ${formattedDate}`,
        ]);
        continue;
      }

      let nameExists = existedNames.find((item) => item.name === finalName);

      if (!nameExists) {
        if (!uniqueNames.has(finalName)) {
          uniqueNames.add(finalName);
          newNamePromises.push(
            Models.Name.create(
              {
                name: finalName,
                category: category,
                userId: req.user.id,
              },
              { transaction }
            )
          );
        }
        processedNames.add(`${finalName}-${formattedDate}`);

        continue;
      }

      const dateExists = await Models.Data.findOne(
        {
          where: { [Op.and]: [{ name: nameExists.id }, { date: body.date }] },
        },
        { transaction }
      );

      if (dateExists) {
        updatedData.push(
          dateExists.update(
            {
              amount: amount,
            },
            { transaction }
          )
        );

        processedNames.add(`${nameExists.name}-${formattedDate}`);
      } else {
        newData.push(
          Models.Data.create(
            {
              name: nameExists.id,
              date: formattedDate,
              amount,
            },
            { transaction }
          )
        );
      }
    }

    const newNames = await Promise.all(newNamePromises)
      .then((result) => {
        return result;
      })
      .catch((error) => {
        errorsInRow.push([`${messages.errorCreatingNames}: ${error.message}`]);
        return [];
      });

    for (const newName of newNames) {
      const row = excelData.find(
        (row) =>
          row.name.toLowerCase().trim() === newName.name.toLowerCase().trim()
      );
      if (row) {
        let { date, amount } = row;
        let formattedDate = parseDate(date);
        formattedDate = new Date(
          formattedDate.getTime() +
            Math.abs(formattedDate.getTimezoneOffset() * 60000)
        );

        newData.push(
          Models.Data.create(
            {
              name: newName.id,
              date: formattedDate,
              amount,
            },
            { transaction }
          )
        );
      }
    }

    await Promise.all(updatedData)
      .then((result) => {
        createdData.push(...result);
      })
      .catch((error) => {
        errorsInRow.push([`${messages.errorUpdatingData}: ${error.message}`]);
      });

    await Promise.all(newData)
      .then((result) => {
        createdData.push(...result);
      })
      .catch((error) => {
        errorsInRow.push([`${messages.errorCreatingData}: ${error.message}`]);
      });

    console.log("Processed Names: ", processedNames);

    if (errorsInRow.length !== 0) {
      hasRolledBack = true;
      await transaction.rollback();
      return errorResponseData(
        res,
        messages.errorInExcelFile,
        errorsInRow,
        400
      );
    } else {
      await transaction.commit();
      return successResponseData(
        res,
        createdData,
        200,
        messages.dataAddedSuccess
      );
    }
  } catch (error) {
    if (!hasRolledBack) {
      await transaction.rollback();
    }
    console.log(error);
    return errorResponseWithoutData(
      res,
      `${messages.somethingWentWrong}: ${error}`,
      400
    );
  }
};

module.exports.getData = async (req, res) => {
  try {
    const validationResult = getDataSchema(req.body, res);

    if (validationResult !== false) return;

    const { startDate, endDate } = req.body;

    console.log(startDate, endDate);

    const data = await Models.Data.findAll({
      where: {
        date: {
          [Op.gte]: new Date(startDate),
          [Op.lte]: new Date(endDate),
        },
      },
    });

    if (!data) {
      return errorResponseWithoutData(res, messages.errorGettingData, 400);
    }

    if (data.length === 0) {
      return successResponseWithoutData(
        res,
        messages.dontHaveDataBetweenDays,
        200
      );
    }

    let filteredData = [];

    let nameTotal = {};

    let total = 0;
    let totalA = 0;
    let totalB = 0;
    let totalC = 0;
    let totalD = 0;
    let totalE = 0;
    for (let i = 0; i < data.length; i++) {
      let name = await Models.Name.findOne({
        where: { id: data[i].dataValues.name },
      });

      filteredData.push({
        name: name.name,
        category: name.category,
        date: data[i].dataValues.date,
        amount: data[i].dataValues.amount,
      });

      if (name.category === "a") {
        totalA += data[i].dataValues.amount;
      }

      if (name.category === "b") {
        totalB += data[i].dataValues.amount;
      }

      if (name.category === "c") {
        totalC += data[i].dataValues.amount;
      }

      if (name.category === "d") {
        totalD += data[i].dataValues.amount;
      }

      if (name.category === "e") {
        totalE += data[i].dataValues.amount;
      }

      total += data[i].dataValues.amount;

      let Name = name.name;

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

    xlsx.writeFile(workbook, "customData.xlsx");

    return successResponseWithoutData(res, messages.processingData, 200);
  } catch (error) {
    console.log(error);
    return errorResponseData(res, messages.somethingWentWrong, error, 400);
  }
};