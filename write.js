const ExcelJS = require('exceljs');
const fs = require('fs');

// Import the data

const data = require('./data.json');

// Create a new workbook

const workbook = new ExcelJS.Workbook();

// Import the template

templateName = 'penyata-akaun-hr-2022.xlsx'

/* 
1. Go to form 1 data
2. Go to part 1 data
3. Go to the first data
4. Get the homeroom
5. Create a copy of the template with the homeroom name
6. Open the copy
7. Write the data to the copy
8. Save the copy
9. Go to the next homeroom
10. Repeat steps 5-9
11. Repeat steps 1-10 for form 2 to 5
*/

// Form 1: 15 homerooms

// For every homeroom, create a copy of the template

function copyFromTemplateForm1() {
    for (let i = 0; i < 15; i++) {
        // Get the homeroom name
        homeroom = data.form1.part1.data[i][1];
    
        // Create a copy of the template with the homeroom name
        fs.copyFile(templateName, `./penyata-akhir/form1/1${homeroom}-2022.xlsx`, (err) => {
            if (err) throw err;
            // console.log(`${homeroom}.xlsx was copied to specified directory`);
        });
    }

    console.log('Copied all homerooms from form 1');
}

copyFromTemplateForm1()

async function writeDataToCopyForm1() {
    for (let i = 0; i < 15; i++) {
        // Get the homeroom name
        homeroom = data.form1.part1.data[i][1];

        // Open the copy

        const workbook = new ExcelJS.Workbook();

        await workbook.xlsx.readFile(`./penyata-akhir/form1/1${homeroom}-2022.xlsx`);

        const worksheet = workbook.getWorksheet('Sheet1');

        // Write the data to the copy

        // Change the homeroom name

        const cell = worksheet.getCell('B8');

        cell.value = `HOMEROOM ${homeroom} 2022`;

        // Part 1

        cellpt1_1 = worksheet.getCell('D16');
        cellpt1_1.value = data.form1.part1.data[i][3];

        cellpt1_2 = worksheet.getCell('E16');
        cellpt1_2.value = data.form1.part1.data[i][4];

        cellpt1_3 = worksheet.getCell('D17');
        cellpt1_3.value = data.form1.part1.data[i][5];

        cellpt1_4 = worksheet.getCell('E17');
        cellpt1_4.value = data.form1.part1.data[i][6];

        cellpt1_5 = worksheet.getCell('D18');
        cellpt1_5.value = data.form1.part1.data[i][7] + data.form1.part1.data[i][9];

        cellpt1_6 = worksheet.getCell('E18');
        cellpt1_6.value = data.form1.part1.data[i][8] + data.form1.part1.data[i][10];
    
        // Part 2

        cellpt2_1 = worksheet.getCell('E21');
        cellpt2_1.value = data.form1.part3.data[i][0] + data.form1.part3.data[i][1] + data.form1.part3.data[i][2];

        cellpt2_2 = worksheet.getCell('E22');
        cellpt2_2.value = data.form1.part3.data[i][3] + data.form1.part3.data[i][4] + data.form1.part3.data[i][5];

        cellpt2_3 = worksheet.getCell('E23');
        cellpt2_3.value = data.form1.part3.data[i][6] + data.form1.part3.data[i][7] + data.form1.part3.data[i][8];

        // Part 3

        cellpt3_1 = worksheet.getCell('D26');
        cellpt3_1.value = data.form1.part2.data[i][0];

        cellpt3_2 = worksheet.getCell('E26');
        cellpt3_2.value = data.form1.part2.data[i][1];

        cellpt3_3 = worksheet.getCell('D27');
        cellpt3_3.value = data.form1.part2.data[i][2];

        cellpt3_4 = worksheet.getCell('E27');
        cellpt3_4.value = data.form1.part2.data[i][3];

        cellpt3_5 = worksheet.getCell('D28');
        cellpt3_5.value = data.form1.part2.data[i][4];

        cellpt3_6 = worksheet.getCell('E28');
        cellpt3_6.value = data.form1.part2.data[i][5];

        cellpt3_7 = worksheet.getCell('D29');
        cellpt3_7.value = data.form1.part2.data[i][6];

        cellpt3_8 = worksheet.getCell('E29');
        cellpt3_8.value = data.form1.part2.data[i][7];

        // Part 4

        cellpt4_1 = worksheet.getCell('D32');
        cellpt4_1.value = data.form1.part4.data[i][0];

        cellpt4_2 = worksheet.getCell('E32');
        cellpt4_2.value = data.form1.part4.data[i][1];

        cellpt4_3 = worksheet.getCell('D33');
        cellpt4_3.value = data.form1.part4.data[i][2];

        cellpt4_4 = worksheet.getCell('E33');
        cellpt4_4.value = data.form1.part4.data[i][3];

        cellpt4_5 = worksheet.getCell('D36');
        cellpt4_5.value = data.form1.part4.data[i][4];

        cellpt4_6 = worksheet.getCell('E36');
        cellpt4_6.value = data.form1.part4.data[i][5];

        cellpt4_7 = worksheet.getCell('D37');
        cellpt4_7.value = data.form1.part4.data[i][6];

        cellpt4_8 = worksheet.getCell('E37');
        cellpt4_8.value = data.form1.part4.data[i][7];

        // Write the data to the copy

        await workbook.xlsx.writeFile(`./penyata-akhir/form1/penyata-akhir-1${homeroom}-2022.xlsx`);

    }
}

writeDataToCopyForm1()

console.log('Done writing data to copies of form 1');

// Form 2

async function copyFromTemplateForm2() {
    for (let i = 0; i < 15; i++) {
        // Get the homeroom name
        homeroom = data.form2.part1.data[i][1];
    
        // Create a copy of the template with the homeroom name
        fs.copyFile(templateName, `./penyata-akhir/form2/2${homeroom}-2022.xlsx`, (err) => {
            if (err) throw err;
            // console.log(`${homeroom}.xlsx was copied to specified directory`);
        });
    }

    console.log('Copied all homerooms from form 2');
}

copyFromTemplateForm2()

async function writeDataToCopyForm2() {
    for (let i = 0; i < 15; i++) {
        // Get the homeroom name
        homeroom = data.form2.part1.data[i][1];

        // Open the copy

        const workbook = new ExcelJS.Workbook();

        await workbook.xlsx.readFile(`./penyata-akhir/form2/2${homeroom}-2022.xlsx`);

        const worksheet = workbook.getWorksheet('Sheet1');

        // Write the data to the copy

        // Change the homeroom name

        const cell = worksheet.getCell('B8');

        cell.value = `HOMEROOM ${homeroom} 2022`;

        // Part 1

        cellpt1_1 = worksheet.getCell('D16');
        cellpt1_1.value = data.form2.part1.data[i][3];

        cellpt1_2 = worksheet.getCell('E16');
        cellpt1_2.value = data.form2.part1.data[i][4];

        cellpt1_3 = worksheet.getCell('D17');
        cellpt1_3.value = data.form2.part1.data[i][5];

        cellpt1_4 = worksheet.getCell('E17');
        cellpt1_4.value = data.form2.part1.data[i][6];

        cellpt1_5 = worksheet.getCell('D18');
        cellpt1_5.value = data.form2.part1.data[i][7] + data.form2.part1.data[i][9];

        cellpt1_6 = worksheet.getCell('E18');
        cellpt1_6.value = data.form2.part1.data[i][8] + data.form2.part1.data[i][10];
    
        // Part 2

        cellpt2_1 = worksheet.getCell('E21');
        cellpt2_1.value = data.form2.part3.data[i][0] + data.form2.part3.data[i][1] + data.form2.part3.data[i][2];

        cellpt2_2 = worksheet.getCell('E22');
        cellpt2_2.value = data.form2.part3.data[i][3] + data.form2.part3.data[i][4] + data.form2.part3.data[i][5];

        cellpt2_3 = worksheet.getCell('E23');
        cellpt2_3.value = data.form2.part3.data[i][6] + data.form2.part3.data[i][7] + data.form2.part3.data[i][8];

        // Part 3

        cellpt3_1 = worksheet.getCell('D26');
        cellpt3_1.value = data.form2.part2.data[i][0];

        cellpt3_2 = worksheet.getCell('E26');
        cellpt3_2.value = data.form2.part2.data[i][1];

        cellpt3_3 = worksheet.getCell('D27');
        cellpt3_3.value = data.form2.part2.data[i][2];

        cellpt3_4 = worksheet.getCell('E27');
        cellpt3_4.value = data.form2.part2.data[i][3];

        cellpt3_5 = worksheet.getCell('D28');
        cellpt3_5.value = data.form2.part2.data[i][4];

        cellpt3_6 = worksheet.getCell('E28');
        cellpt3_6.value = data.form2.part2.data[i][5];

        cellpt3_7 = worksheet.getCell('D29');
        cellpt3_7.value = data.form2.part2.data[i][6];

        cellpt3_8 = worksheet.getCell('E29');
        cellpt3_8.value = data.form2.part2.data[i][7];

        // Part 4

        cell



