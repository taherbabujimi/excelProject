const Models = require("../models/index");
const { sequelize } = require("../models/index");
const {
  addDataSchema,
  getDataSchema,
  formulaSchema,
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

      const allowedValues = ["name", "category", "date", "amount", "synonym"];

      if (keys.some((value) => !allowedValues.includes(value))) {
        return errorResponseWithoutData(
          res,
          messages.provideAllowedColumns,
          400
        );
      }

      console.log("row: ", row);

      let { name, category, date, amount, synonym } = row;
      if (
        [name, category, date, amount, synonym].some(
          (field) => field === undefined
        )
      ) {
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

      const body = {
        name: finalName,
        category,
        date: formattedDate,
        amount,
        synonym,
      };

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
                synonym: synonym,
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
              createdBy: req.user.id,
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
        console.log(error);
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
              createdBy: req.user.id,
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
          [Op.gte]: new Date(startDate),
          [Op.lte]: new Date(endDate),
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

    const total = data.reduce((sum, item) => sum + item.amount, 0);

    let nameTotal = {};

    filteredData.map((item) => {
      if (Object.hasOwn(nameTotal, `Total ${item.name}`)) {
        nameTotal[`Total ${item.name}`] += item.amount;
      } else {
        nameTotal[`Total ${item.name}`] = item.amount;
      }
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
      ["Grand Total", total.toFixed(2)],
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

    xlsx.writeFile(workbook, "customData.xlsx");

    return successResponseWithoutData(res, messages.processingData, 200);
  } catch (error) {
    console.log(error);
    return errorResponseData(res, messages.somethingWentWrong, error, 400);
  }
};

module.exports.addFormula = async (req, res) => {
  try {
    const validationResponse = formulaSchema(req.body, res);

    if (validationResponse !== false) return;

    const { formula } = req.body;

    let formulaArray = formula.split(" ");
    console.log("formula array : ", formulaArray);

    // let answer = "5A";
    let filteredNames = formulaArray.map((item) => {
      return item.replace(/[^a-z]/gi, "");
    });

    const regex = /:\s*(\d{4}(?:-\d{4})?(?:-\d{4})?|\d{4})/g;
    const matches = [];
    let match;
    while ((match = regex.exec(formula)) !== null) {
      matches.push(match[1]);
    }

    console.log(matches);

    // const input = "key1: value1, key2: value2, key3: value3";
    // const matches = input
    //   .match(/:\s*([^,]+)/g)
    //   ?.map((match) => match.slice(1).trim());

    // console.log("matches: ", matches)

    console.log("formulaname: ", filteredNames);

    let existedNames = [];
    let j = 0;
    for (let i = 0; i <= filteredNames.length; i++) {
      if (filteredNames[i] !== "" && filteredNames[i] !== undefined) {
        console.log(filteredNames[i]);
        existedNames.push(
          Models.Name.findOne({
            where: { synonym: filteredNames[i] },
          })
        );
      }
    }

    let existedNamesResult = [];
    await Promise.all(existedNames)
      .then((result) => {
        // console.log(result);
        existedNamesResult = result;
      })
      .catch((error) => {
        console.log(error);
      });

    // console.log("Existed names result: ", existedNamesResult);

    if (existedNamesResult.includes(null)) {
      return errorResponseWithoutData(
        res,
        "Please provide valid synonyms in the formula",
        400
      );
    }

    const existedUserFormula = await Models.Formula.count({
      where: {
        createdBy: req.user.id,
      },
    });

    if (existedUserFormula === 5) {
      return errorResponseWithoutData(res, messages.Formula5NotAllowed, 400);
    }

    console.log(existedUserFormula);

    const addFormula = await Models.Formula.create({
      formula: formula,
      createdBy: req.user.id,
    });

    if (!addFormula) {
      return errorResponseWithoutData(res, messages.errorAddingFormula);
    }

    return successResponseData(
      res,
      addFormula,
      200,
      messages.successAddingFormula
    );
  } catch (error) {
    console.log(error);
    return errorResponseData(res, messages.somethingWentWrong, error, 400);
  }
};
