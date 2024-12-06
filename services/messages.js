const { invalid } = require("joi");

module.exports = {
  messages: {
    error: "Something went wrong.",
    errorWhileValidating: "Error while validating values.",
    userAlreadyExists: "User with this email already exists.",
    userCreatedSuccessfully: "User Created Successfully.",
    userLoginSuccess: "User logged in successfully.",
    userNotExists: "User with this email does not exists.",
    somethingWentWrong: "Something went wrong.",
    incorrectCredentials: "Please provide valid user credentials.",
    invalidToken: "Invalid Access Token.",
    badRequest: "Unauthorized request.",
    nameAlreadyExists: "Name Already Exists.",
    nameAddedSuccess: "Name added successfully.",
    dataAddedSuccess: "Data added successfullu from excel file.",
    userNotHaveThisName: "You don't have this name, Please add the name first.",
    dateAlreadyExists: "Data with this name and date already exists.",
    invalidColumnNames:
      "Please provide valid column names inside the excel file",
    errorGettingData: "Error while fetching data.",
    dontHaveDataBetweenDays: "You don't have data between this dates.",
    processingData: "Data is bieng proccessed.",
    excelFileEmpty: "Excel file is empty, Please add some data first.",
    provideAllowedColumns: "Please provide column names which are allowed.",
    provideFieldAndColumnProperly:
      "Please provide each field and column properly.",
    errorCreatingNames: "Error creating new names",
    errorUpdatingData: "Error updating data",
    errorCreatingData: "Error creating new data",
    errorInExcelFile: "There is error inside Excel file.",
    errorAddingFormula: "Something went wrong while adding formula.",
    successAddingFormula: "Formula added successfully.",
    Formula5NotAllowed: "More than 5 formulas are not allowed",
    futureDateNotAllowed: "Dates in future are not allowed.",
    formulaNameAlreadyExists: "Formula with this name already exists.",
    improperBrackets: "Please provide proper opening and closing brackets.",
    emptyBracketsNotAllowed:
      "Empty brackets inside the formula is not allowed.",
    invalidExpressions: "Please remove the following invalid expressions.",
    invalidFormula:
      "Provided formula is invalid. Please provide valid formula.",
  },
};
