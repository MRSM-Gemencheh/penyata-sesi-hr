const ExcelJS = require('exceljs');
const fs = require('fs');

// Import the data

const data = require('./data.json');

// Create a new workbook

const workbook = new ExcelJS.Workbook();

// Import the template

filename = 'penyata-akaun-hr-2022.xlsx'

async function readExcelFileFromSystem(filename) {

    await workbook.xlsx.readFile(filename);

}

readExcelFileFromSystem(filename)

console.log(data)