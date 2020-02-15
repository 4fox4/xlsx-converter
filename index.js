"use strict";

const inquirer = require("inquirer");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const { format } = require("date-fns");

const inputPath = path.resolve(__dirname, "input");
const outputPath = path.resolve(__dirname, "output");

const filesList = fs.readdirSync(inputPath);
// console.log(filesList);
const extensionsList = [
  "json",
  "csv",
  "txt",
  "html",
  "formulae",
];

inquirer
  .prompt([
    {
      type: "list",
      name: "file",
      message: "Select the input file",
      choices: filesList
    },
    {
      type: "list",
      name: "extension",
      message: "Select the output format",
      choices: extensionsList,
    },
  ])
  .then(answers => {
    // console.log(JSON.stringify(answers, null, 2));
    const inputFileName = answers.file;
    const outputExtension = answers.extension;
    console.log("TCL: outputExtension", outputExtension)
    const inputFilePath = path.resolve(inputPath, inputFileName);
    // console.log("TCL: inputFilePath", inputFilePath)
    const workbook = XLSX.readFile(inputFilePath, { type: "buffer" });
    // const buf = fs.readFileSync(inputFilePath);
    // const workbook = XLSX.read(buf, { type: "buffer" });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    console.log("TCL: firstSheetName", firstSheetName);
    const output = XLSX.utils[`sheet_to_${outputExtension}`](worksheet);
    // console.log("TCL: output", output);
    const newFileName = inputFileName
      .split(".")
      .slice(0, -1)
      .join("_");
    const currentDate = format(new Date(), "ddMMyyyy-HHmm");
    const outputFileName = `${newFileName}_${currentDate}.${outputExtension}`;
    console.log("TCL: outputFileName", outputFileName);
    const outputFilePath = path.resolve(outputPath, outputFileName);
    const outputData = typeof output === "object" ? JSON.stringify(outputData, null, 2) : output;
    fs.writeFileSync(outputFilePath, outputData)
  });
