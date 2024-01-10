const ExcelJS = require('exceljs');
const fs = require('fs');
const prompt = require('prompt-sync')({ sigint: true });

// Get what year they want to generate the penyata akhir for
year = prompt("Masukkan tahun: (20XX) ")

// Import the data in the raw_data folder depending on the year
const data = require(`./data_json/merit_demerit_data_${year}.json`)

// Create a new workbook
const workbook = new ExcelJS.Workbook();

// Import the template for the penyata akhir
templateName = `./src/Template_Penyata_Akaun_HR_${year}.xlsx`

// For every homeroom, create a copy of the template
function copyFromTemplateForm1() {
    for (let i = 0; i < 15; i++) {
        // Get the homeroom name
        homeroom = data.form1.part1.data[i][1];

        // Create a copy of the template with the homeroom name
        fs.copyFile(templateName, `./penyata-akhir-${year}/form1/1${homeroom}-${year}.xlsx`, (err) => {
            if (err) throw err;
            // console.log(`${homeroom}.xlsx was copied to specified directory`);
        });
    }

    console.log('Copied all homerooms from form 1');
}

async function writeDataToCopyForm1() {
    for (let i = 0; i < 15; i++) {
        // Get the homeroom name
        homeroom = data.form1.part1.data[i][1];

        // Open the copy

        const workbook = new ExcelJS.Workbook();

        await workbook.xlsx.readFile(`./penyata-akhir-${year}/form1/1${homeroom}-${year}.xlsx`);

        const worksheet = workbook.getWorksheet('Sheet1');

        // Write the data to the copy

        // Change the homeroom name

        const cell = worksheet.getCell('B8');

        cell.value = `HOMEROOM 1${homeroom} ${year}`;

        // Part 1

        cellpt1_1 = worksheet.getCell('D16');
        cellpt1_1.value = data.form1.part1.data[i][3] || 0;

        cellpt1_2 = worksheet.getCell('E16');
        cellpt1_2.value = data.form1.part1.data[i][4] || 0;

        cellpt1_3 = worksheet.getCell('D17');
        cellpt1_3.value = data.form1.part1.data[i][5] || 0;

        cellpt1_4 = worksheet.getCell('E17');
        cellpt1_4.value = data.form1.part1.data[i][6] || 0;

        cellpt1_5 = worksheet.getCell('D18');
        cellpt1_5.value = data.form1.part1.data[i][7] || 0;

        cellpt1_6 = worksheet.getCell('E18');
        cellpt1_6.value = data.form1.part1.data[i][8] || 0;

        cellpt1_7 = worksheet.getCell('D19');
        cellpt1_7.value = data.form1.part1.data[i][9] || 0;

        cellpt1_8 = worksheet.getCell('E19');
        cellpt1_8.value = data.form1.part1.data[i][10] || 0;

        // Part 2 (Cells D22, E22, D23, E23, D24, E24, D25, E25)

        cellpt2_1 = worksheet.getCell('D22');
        cellpt2_1.value = data.form1.part2.data[i][0] || 0;

        cellpt2_2 = worksheet.getCell('E22');
        cellpt2_2.value = data.form1.part2.data[i][1] || 0;

        cellpt2_3 = worksheet.getCell('D23');
        cellpt2_3.value = data.form1.part2.data[i][2] || 0;

        cellpt2_4 = worksheet.getCell('E23');
        cellpt2_4.value = data.form1.part2.data[i][3] || 0;

        cellpt2_5 = worksheet.getCell('D24');
        cellpt2_5.value = data.form1.part2.data[i][4] || 0;

        cellpt2_6 = worksheet.getCell('E24');
        cellpt2_6.value = data.form1.part2.data[i][5] || 0;

        cellpt2_7 = worksheet.getCell('D25');
        cellpt2_7.value = data.form1.part2.data[i][6] || 0;

        cellpt2_8 = worksheet.getCell('E25');
        cellpt2_8.value = data.form1.part2.data[i][7] || 0;

        // Part 3 (Cells E28, E29, E30, E31, )

        cellpt3_1 = worksheet.getCell('E28');
        cellpt3_1.value = data.form1.part3.data[i][0] + data.form1.part3.data[i][1] + data.form1.part3.data[i][2] || 0;

        cellpt3_2 = worksheet.getCell('E29');
        cellpt3_2.value = data.form1.part3.data[i][3] + data.form1.part3.data[i][4] + data.form1.part3.data[i][5] || 0;

        cellpt3_3 = worksheet.getCell('E30');
        cellpt3_3.value = data.form1.part3.data[i][6] + data.form1.part3.data[i][7] + data.form1.part3.data[i][8] || 0;

        cellpt3_4 = worksheet.getCell('E31');
        cellpt3_4.value = 0

        // Part 4 (Cells D34, E34, D35, E35, D36, E36, D37, E37, D38, E38, D39, E39, D40, E40) depending on how many pertandingans there are

        // Check number of pertandingans and only write to their respective cells, going down for each pertandingan

        cellpt4_1 = worksheet.getCell('D34');
        cellpt4_1.value = data.form1.part4.data[i + 1][0] || 0;

        cellpt4_2 = worksheet.getCell('E34');
        cellpt4_2.value = data.form1.part4.data[i + 1][1] || 0;

        cellpt4_3 = worksheet.getCell('D35');
        cellpt4_3.value = data.form1.part4.data[i + 1][2] || 0;

        cellpt4_4 = worksheet.getCell('E35');
        cellpt4_4.value = data.form1.part4.data[i + 1][3] || 0;

        cellpt4_5 = worksheet.getCell('D36');
        cellpt4_5.value = data.form1.part4.data[i + 1][4] || 0;

        cellpt4_6 = worksheet.getCell('E36');
        cellpt4_6.value = data.form1.part4.data[i + 1][5] || 0;

        cellpt4_7 = worksheet.getCell('D37');
        cellpt4_7.value = data.form1.part4.data[i + 1][6] || 0;

        cellpt4_8 = worksheet.getCell('E37');
        cellpt4_8.value = data.form1.part4.data[i + 1][7] || 0;

        cellpt4_9 = worksheet.getCell('D38');
        cellpt4_9.value = data.form1.part4.data[i + 1][8] || 0;

        cellpt4_10 = worksheet.getCell('E38');
        cellpt4_10.value = data.form1.part4.data[i + 1][9] || 0;

        cellpt4_11 = worksheet.getCell('D39');
        cellpt4_11.value = data.form1.part4.data[i + 1][10] || 0;

        cellpt4_12 = worksheet.getCell('E39');
        cellpt4_12.value = data.form1.part4.data[i + 1][11] || 0;

        cellpt4_13 = worksheet.getCell('D40');
        cellpt4_13.value = data.form1.part4.data[i + 1][12] || 0;

        cellpt4_14 = worksheet.getCell('E40');
        cellpt4_14.value = data.form1.part4.data[i + 1][13] || 0;

        cellpt4_15 = worksheet.getCell('D41');
        cellpt4_15.value = data.form1.part4.data[i + 1][14] || 0;

        cellpt4_16 = worksheet.getCell('E41');
        cellpt4_16.value = data.form1.part4.data[i + 1][15] || 0;

        form1PertandinganNames = data.form1.part4.data[0];

        // Fil in the pertandingan names from C34 until C40

        for (let j = 0; j < form1PertandinganNames.length; j++) {
            cellpt4_9 = worksheet.getCell(`C${34 + j}`);
            cellpt4_9.value = form1PertandinganNames[j];

            cellpt4_10 = worksheet.getCell('C41');
            cellpt4_10.value = '';
        }

        // Write the data to the copy

        await workbook.xlsx.writeFile(`./penyata-akhir-${year}/form1/1${homeroom}-${year}.xlsx`);

    }

    console.log('Done writing data to copies of Form 1');

}


// Form 2

async function copyFromTemplateForm2() {
    for (let i = 0; i < 18; i++) {
        // Get the homeroom name
        homeroom = data.form2.part1.data[i][1];

        // Create a copy of the template with the homeroom name
        fs.copyFile(templateName, `./penyata-akhir-${year}/form2/2${homeroom}-${year}.xlsx`, (err) => {
            if (err) throw err;
            // console.log(`${homeroom}.xlsx was copied to specified directory`);
        });
    }

    console.log('Copied all homerooms from form 2');
}

async function writeDataToCopyForm2() {
    for (let i = 0; i < 18; i++) {
        // Get the homeroom name
        homeroom = data.form2.part1.data[i][1];

        // Open the copy

        const workbook = new ExcelJS.Workbook();

        await workbook.xlsx.readFile(`./penyata-akhir-${year}/form2/2${homeroom}-${year}.xlsx`);

        const worksheet = workbook.getWorksheet('Sheet1');

        // Write the data to the copy

        // Change the homeroom name

        const cell = worksheet.getCell('B8');

        cell.value = `HOMEROOM 2${homeroom} ${year}`;

        // Part 1

        cellpt1_1 = worksheet.getCell('D16');
        cellpt1_1.value = data.form2.part1.data[i][3] || 0;

        cellpt1_2 = worksheet.getCell('E16');
        cellpt1_2.value = data.form2.part1.data[i][4] || 0;

        cellpt1_3 = worksheet.getCell('D17');
        cellpt1_3.value = data.form2.part1.data[i][5] || 0;

        cellpt1_4 = worksheet.getCell('E17');
        cellpt1_4.value = data.form2.part1.data[i][6] || 0;

        cellpt1_5 = worksheet.getCell('D18');
        cellpt1_5.value = data.form2.part1.data[i][7] || 0;

        cellpt1_6 = worksheet.getCell('E18');
        cellpt1_6.value = data.form2.part1.data[i][8] || 0;

        cellpt1_7 = worksheet.getCell('D19');
        cellpt1_7.value = data.form2.part1.data[i][9] || 0;

        cellpt1_8 = worksheet.getCell('E19');
        cellpt1_8.value = data.form2.part1.data[i][10] || 0;

        // Part 2 (Cells D22, E22, D23, E23, D24, E24, D25, E25)

        cellpt2_1 = worksheet.getCell('D22');
        cellpt2_1.value = data.form2.part2.data[i][0] || 0;

        cellpt2_2 = worksheet.getCell('E22');
        cellpt2_2.value = data.form2.part2.data[i][1] || 0;

        cellpt2_3 = worksheet.getCell('D23');
        cellpt2_3.value = data.form2.part2.data[i][2] || 0;

        cellpt2_4 = worksheet.getCell('E23');
        cellpt2_4.value = data.form2.part2.data[i][3] || 0;

        cellpt2_5 = worksheet.getCell('D24');
        cellpt2_5.value = data.form2.part2.data[i][4] || 0;

        cellpt2_6 = worksheet.getCell('E24');
        cellpt2_6.value = data.form2.part2.data[i][5] || 0;

        cellpt2_7 = worksheet.getCell('D25');
        cellpt2_7.value = data.form2.part2.data[i][6] || 0;

        cellpt2_8 = worksheet.getCell('E25');
        cellpt2_8.value = data.form2.part2.data[i][7] || 0;

        // Part 3 (Cells E28, E29, E30, E31, )

        cellpt3_1 = worksheet.getCell('E28');
        cellpt3_1.value = data.form2.part3.data[i][0] + data.form2.part3.data[i][1] + data.form2.part3.data[i][2] || 0;

        cellpt3_2 = worksheet.getCell('E29');
        cellpt3_2.value = data.form2.part3.data[i][3] + data.form2.part3.data[i][4] + data.form2.part3.data[i][5] || 0;

        cellpt3_3 = worksheet.getCell('E30');
        cellpt3_3.value = data.form2.part3.data[i][6] + data.form2.part3.data[i][7] + data.form2.part3.data[i][8] || 0;

        cellpt3_4 = worksheet.getCell('E31');
        cellpt3_4.value = 0

        // Part 4 (Cells D34, E34, D35, E35, D36, E36, D37, E37, D38, E38, D39, E39, D40, E40) depending on how many pertandingans there are

        // Check number of pertandingans and only write to their respective cells, going down for each pertandingan

        cellpt4_1 = worksheet.getCell('D34');
        cellpt4_1.value = data.form2.part4.data[i + 1][0] || 0;

        cellpt4_2 = worksheet.getCell('E34');
        cellpt4_2.value = data.form2.part4.data[i + 1][1] || 0;

        cellpt4_3 = worksheet.getCell('D35');
        cellpt4_3.value = data.form2.part4.data[i + 1][2] || 0;

        cellpt4_4 = worksheet.getCell('E35');
        cellpt4_4.value = data.form2.part4.data[i + 1][3] || 0;

        cellpt4_5 = worksheet.getCell('D36');
        cellpt4_5.value = data.form2.part4.data[i + 1][4] || 0;

        cellpt4_6 = worksheet.getCell('E36');
        cellpt4_6.value = data.form2.part4.data[i + 1][5] || 0;

        cellpt4_7 = worksheet.getCell('D37');
        cellpt4_7.value = data.form2.part4.data[i + 1][6] || 0;

        cellpt4_8 = worksheet.getCell('E37');
        cellpt4_8.value = data.form2.part4.data[i + 1][7] || 0;

        cellpt4_9 = worksheet.getCell('D38');
        cellpt4_9.value = data.form2.part4.data[i + 1][8] || 0;

        cellpt4_10 = worksheet.getCell('E38');
        cellpt4_10.value = data.form2.part4.data[i + 1][9] || 0;

        cellpt4_11 = worksheet.getCell('D39');
        cellpt4_11.value = data.form2.part4.data[i + 1][10] || 0;

        cellpt4_12 = worksheet.getCell('E39');
        cellpt4_12.value = data.form2.part4.data[i + 1][11] || 0;

        cellpt4_13 = worksheet.getCell('D40');
        cellpt4_13.value = data.form2.part4.data[i + 1][12] || 0;

        cellpt4_14 = worksheet.getCell('E40');
        cellpt4_14.value = data.form2.part4.data[i + 1][13] || 0;

        cellpt4_15 = worksheet.getCell('D41');
        cellpt4_15.value = data.form2.part4.data[i + 1][14] || 0;

        cellpt4_16 = worksheet.getCell('E41');
        cellpt4_16.value = data.form2.part4.data[i + 1][15] || 0;

        form2PertandinganNames = data.form2.part4.data[0];

        // Fil in the pertandingan names from C34 until C40

        for (let j = 0; j < form2PertandinganNames.length; j++) {
            cellpt4_9 = worksheet.getCell(`C${34 + j}`);
            cellpt4_9.value = form2PertandinganNames[j];

            cellpt4_10 = worksheet.getCell('C41');
            cellpt4_10.value = '';
        }

        // Write the data to the copy

        await workbook.xlsx.writeFile(`./penyata-akhir-${year}/form2/2${homeroom}-${year}.xlsx`);



    }

    console.log('Done writing data to copies of Form 2')


}

// Form 3

async function copyFromTemplateForm3() {
    for (let i = 0; i < 18; i++) {
        // Get the homeroom name
        homeroom = data.form3.part1.data[i][1];

        // Create a copy of the template with the homeroom name
        fs.copyFile(templateName, `./penyata-akhir-${year}/form3/3${homeroom}-${year}.xlsx`, (err) => {
            if (err) throw err;
            // console.log(`${homeroom}.xlsx was copied to specified directory`);
        });
    }

    console.log('Copied all homerooms from form 3');
}

async function writeDataToCopyForm3() {
    for (let i = 0; i < 18; i++) {
        // Get the homeroom name
        homeroom = data.form3.part1.data[i][1];

        // Open the copy

        const workbook = new ExcelJS.Workbook();

        await workbook.xlsx.readFile(`./penyata-akhir-${year}/form3/3${homeroom}-${year}.xlsx`);

        const worksheet = workbook.getWorksheet('Sheet1');

        // Write the data to the copy

        // Change the homeroom name

        const cell = worksheet.getCell('B8');

        cell.value = `HOMEROOM 3${homeroom} ${year}`;

        // Part 1

        cellpt1_1 = worksheet.getCell('D16');
        cellpt1_1.value = data.form3.part1.data[i][3] || 0;

        cellpt1_2 = worksheet.getCell('E16');
        cellpt1_2.value = data.form3.part1.data[i][4] || 0;

        cellpt1_3 = worksheet.getCell('D17');
        cellpt1_3.value = data.form3.part1.data[i][5] || 0;

        cellpt1_4 = worksheet.getCell('E17');
        cellpt1_4.value = data.form3.part1.data[i][6] || 0;

        cellpt1_5 = worksheet.getCell('D18');
        cellpt1_5.value = data.form3.part1.data[i][7] || 0;

        cellpt1_6 = worksheet.getCell('E18');
        cellpt1_6.value = data.form3.part1.data[i][8] || 0;

        cellpt1_7 = worksheet.getCell('D19');
        cellpt1_7.value = data.form3.part1.data[i][9] || 0;

        cellpt1_8 = worksheet.getCell('E19');
        cellpt1_8.value = data.form3.part1.data[i][10] || 0;

        // Part 2 (Cells D22, E22, D23, E23, D24, E24, D25, E25)

        cellpt2_1 = worksheet.getCell('D22');
        cellpt2_1.value = data.form3.part2.data[i][0] || 0;

        cellpt2_2 = worksheet.getCell('E22');
        cellpt2_2.value = data.form3.part2.data[i][1] || 0;

        cellpt2_3 = worksheet.getCell('D23');
        cellpt2_3.value = data.form3.part2.data[i][2] || 0;

        cellpt2_4 = worksheet.getCell('E23');
        cellpt2_4.value = data.form3.part2.data[i][3] || 0;

        cellpt2_5 = worksheet.getCell('D24');
        cellpt2_5.value = data.form3.part2.data[i][4] || 0;

        cellpt2_6 = worksheet.getCell('E24');
        cellpt2_6.value = data.form3.part2.data[i][5] || 0;

        cellpt2_7 = worksheet.getCell('D25');
        cellpt2_7.value = data.form3.part2.data[i][6] || 0;

        cellpt2_8 = worksheet.getCell('E25');
        cellpt2_8.value = data.form3.part2.data[i][7] || 0;

        // Part 3 (Cells E28, E29, E30, E31, )

        cellpt3_1 = worksheet.getCell('E28');
        cellpt3_1.value = data.form3.part3.data[i][0] + data.form3.part3.data[i][1] + data.form3.part3.data[i][2] || 0;

        cellpt3_2 = worksheet.getCell('E29');
        cellpt3_2.value = data.form3.part3.data[i][3] + data.form3.part3.data[i][4] + data.form3.part3.data[i][5] || 0;

        cellpt3_3 = worksheet.getCell('E30');
        cellpt3_3.value = data.form3.part3.data[i][6] + data.form3.part3.data[i][7] + data.form3.part3.data[i][8] || 0;

        cellpt3_4 = worksheet.getCell('E31');
        cellpt3_4.value = 0

        // Part 4 (Cells D34, E34, D35, E35, D36, E36, D37, E37, D38, E38, D39, E39, D40, E40) depending on how many pertandingans there are

        // Check number of pertandingans and only write to their respective cells, going down for each pertandingan

        cellpt4_1 = worksheet.getCell('D34');
        cellpt4_1.value = data.form3.part4.data[i + 1][0] || 0;

        cellpt4_2 = worksheet.getCell('E34');
        cellpt4_2.value = data.form3.part4.data[i + 1][1] || 0;

        cellpt4_3 = worksheet.getCell('D35');
        cellpt4_3.value = data.form3.part4.data[i + 1][2] || 0;

        cellpt4_4 = worksheet.getCell('E35');
        cellpt4_4.value = data.form3.part4.data[i + 1][3] || 0;

        cellpt4_5 = worksheet.getCell('D36');
        cellpt4_5.value = data.form3.part4.data[i + 1][4] || 0;

        cellpt4_6 = worksheet.getCell('E36');
        cellpt4_6.value = data.form3.part4.data[i + 1][5] || 0;

        cellpt4_7 = worksheet.getCell('D37');
        cellpt4_7.value = data.form3.part4.data[i + 1][6] || 0;

        cellpt4_8 = worksheet.getCell('E37');
        cellpt4_8.value = data.form3.part4.data[i + 1][7] || 0;

        cellpt4_9 = worksheet.getCell('D38');
        cellpt4_9.value = data.form3.part4.data[i + 1][8] || 0;

        cellpt4_10 = worksheet.getCell('E38');
        cellpt4_10.value = data.form3.part4.data[i + 1][9] || 0;

        cellpt4_11 = worksheet.getCell('D39');
        cellpt4_11.value = data.form3.part4.data[i + 1][10] || 0;

        cellpt4_12 = worksheet.getCell('E39');
        cellpt4_12.value = data.form3.part4.data[i + 1][11] || 0;

        cellpt4_13 = worksheet.getCell('D40');
        cellpt4_13.value = data.form3.part4.data[i + 1][12] || 0;

        cellpt4_14 = worksheet.getCell('E40');
        cellpt4_14.value = data.form3.part4.data[i + 1][13] || 0;

        cellpt4_15 = worksheet.getCell('D41');
        cellpt4_15.value = data.form3.part4.data[i + 1][14] || 0;

        cellpt4_16 = worksheet.getCell('E41');
        cellpt4_16.value = data.form3.part4.data[i + 1][15] || 0;

        form3PertandinganNames = data.form3.part4.data[0];

        // Fil in the pertandingan names from C34 until C40

        for (let j = 0; j < form3PertandinganNames.length; j++) {
            cellpt4_9 = worksheet.getCell(`C${34 + j}`);
            cellpt4_9.value = form3PertandinganNames[j];

            cellpt4_10 = worksheet.getCell('C41');
            cellpt4_10.value = '';
        }

        // Write the data to the copy

        await workbook.xlsx.writeFile(`./penyata-akhir-${year}/form3/3${homeroom}-${year}.xlsx`);

    }

    console.log('Done writing data to copies of Form 3')

}

// Form 4

async function copyFromTemplateForm4() {
    for (let i = 0; i < 15; i++) {
        // Get the homeroom name
        homeroom = data.form4.part1.data[i][1];

        // Create a copy of the template with the homeroom name
        fs.copyFile(templateName, `./penyata-akhir-${year}/form4/4${homeroom}-${year}.xlsx`, (err) => {
            if (err) throw err;
            // console.log(`${homeroom}.xlsx was copied to specified directory`);
        });
    }

    console.log('Copied all homerooms from form 4');
}

async function writeDataToCopyForm4() {
    for (let i = 0; i < 15; i++) {
        // Get the homeroom name
        homeroom = data.form4.part1.data[i][1];

        // Open the copy

        const workbook = new ExcelJS.Workbook();

        await workbook.xlsx.readFile(`./penyata-akhir-${year}/form4/4${homeroom}-${year}.xlsx`);

        const worksheet = workbook.getWorksheet('Sheet1');

        // Write the data to the copy

        // Change the homeroom name

        const cell = worksheet.getCell('B8');

        cell.value = `HOMEROOM 4${homeroom} ${year}`;

        // Part 1

        cellpt1_1 = worksheet.getCell('D16');
        cellpt1_1.value = data.form4.part1.data[i][3] || 0;

        cellpt1_2 = worksheet.getCell('E16');
        cellpt1_2.value = data.form4.part1.data[i][4] || 0;

        cellpt1_3 = worksheet.getCell('D17');
        cellpt1_3.value = data.form4.part1.data[i][5] || 0;

        cellpt1_4 = worksheet.getCell('E17');
        cellpt1_4.value = data.form4.part1.data[i][6] || 0;

        cellpt1_5 = worksheet.getCell('D18');
        cellpt1_5.value = data.form4.part1.data[i][7] || 0;

        cellpt1_6 = worksheet.getCell('E18');
        cellpt1_6.value = data.form4.part1.data[i][8] || 0;

        cellpt1_7 = worksheet.getCell('D19');
        cellpt1_7.value = data.form4.part1.data[i][9] || 0;

        cellpt1_8 = worksheet.getCell('E19');
        cellpt1_8.value = data.form4.part1.data[i][10] || 0;

        // Part 2 (Cells D22, E22, D23, E23, D24, E24, D25, E25)

        cellpt2_1 = worksheet.getCell('D22');
        cellpt2_1.value = data.form4.part2.data[i][0] || 0;

        cellpt2_2 = worksheet.getCell('E22');
        cellpt2_2.value = data.form4.part2.data[i][1] || 0;

        cellpt2_3 = worksheet.getCell('D23');
        cellpt2_3.value = data.form4.part2.data[i][2] || 0;

        cellpt2_4 = worksheet.getCell('E23');
        cellpt2_4.value = data.form4.part2.data[i][3] || 0;

        cellpt2_5 = worksheet.getCell('D24');
        cellpt2_5.value = data.form4.part2.data[i][4] || 0;

        cellpt2_6 = worksheet.getCell('E24');
        cellpt2_6.value = data.form4.part2.data[i][5] || 0;

        cellpt2_7 = worksheet.getCell('D25');
        cellpt2_7.value = data.form4.part2.data[i][6] || 0;

        cellpt2_8 = worksheet.getCell('E25');
        cellpt2_8.value = data.form4.part2.data[i][7] || 0;

        // Part 3 (Cells E28, E29, E30, E31, )

        cellpt3_1 = worksheet.getCell('E28');
        cellpt3_1.value = data.form4.part3.data[i][0] + data.form4.part3.data[i][1] + data.form4.part3.data[i][2] || 0;

        cellpt3_2 = worksheet.getCell('E29');
        cellpt3_2.value = data.form4.part3.data[i][3] + data.form4.part3.data[i][4] + data.form4.part3.data[i][5] || 0;

        cellpt3_3 = worksheet.getCell('E30');
        cellpt3_3.value = data.form4.part3.data[i][6] + data.form4.part3.data[i][7] + data.form4.part3.data[i][8] || 0;

        cellpt3_4 = worksheet.getCell('E31');
        cellpt3_4.value = 0

        // Part 4 (Cells D34, E34, D35, E35, D36, E36, D37, E37, D38, E38, D39, E39, D40, E40) depending on how many pertandingans there are

        // Check number of pertandingans and only write to their respective cells, going down for each pertandingan

        cellpt4_1 = worksheet.getCell('D34');
        cellpt4_1.value = data.form4.part4.data[i + 1][0] || 0;

        cellpt4_2 = worksheet.getCell('E34');
        cellpt4_2.value = data.form4.part4.data[i + 1][1] || 0;

        cellpt4_3 = worksheet.getCell('D35');
        cellpt4_3.value = data.form4.part4.data[i + 1][2] || 0;

        cellpt4_4 = worksheet.getCell('E35');
        cellpt4_4.value = data.form4.part4.data[i + 1][3] || 0;

        cellpt4_5 = worksheet.getCell('D36');
        cellpt4_5.value = data.form4.part4.data[i + 1][4] || 0;

        cellpt4_6 = worksheet.getCell('E36');
        cellpt4_6.value = data.form4.part4.data[i + 1][5] || 0;

        cellpt4_7 = worksheet.getCell('D37');
        cellpt4_7.value = data.form4.part4.data[i + 1][6] || 0;

        cellpt4_8 = worksheet.getCell('E37');
        cellpt4_8.value = data.form4.part4.data[i + 1][7] || 0;

        cellpt4_9 = worksheet.getCell('D38');
        cellpt4_9.value = data.form4.part4.data[i + 1][8] || 0;

        cellpt4_10 = worksheet.getCell('E38');
        cellpt4_10.value = data.form4.part4.data[i + 1][9] || 0;

        cellpt4_11 = worksheet.getCell('D39');
        cellpt4_11.value = data.form4.part4.data[i + 1][10] || 0;

        cellpt4_12 = worksheet.getCell('E39');
        cellpt4_12.value = data.form4.part4.data[i + 1][11] || 0;

        cellpt4_13 = worksheet.getCell('D40');
        cellpt4_13.value = data.form4.part4.data[i + 1][12] || 0;

        cellpt4_14 = worksheet.getCell('E40');
        cellpt4_14.value = data.form4.part4.data[i + 1][13] || 0;

        cellpt4_15 = worksheet.getCell('D41');
        cellpt4_15.value = data.form4.part4.data[i + 1][14] || 0;

        cellpt4_16 = worksheet.getCell('E41');
        cellpt4_16.value = data.form4.part4.data[i + 1][15] || 0;

        form4PertandinganNames = data.form4.part4.data[0];

        // Fil in the pertandingan names from C34 until C40

        for (let j = 0; j < form4PertandinganNames.length; j++) {
            cellpt4_9 = worksheet.getCell(`C${34 + j}`);
            cellpt4_9.value = form4PertandinganNames[j];

            cellpt4_10 = worksheet.getCell('C41');
            cellpt4_10.value = '';
        }

        // Write the data to the copy

        await workbook.xlsx.writeFile(`./penyata-akhir-${year}/form4/4${homeroom}-${year}.xlsx`);

    }

    console.log('Done writing data to copies of Form 4')

}


// Form 5

async function copyFromTemplateForm5() {
    for (let i = 0; i < 18; i++) {
        // Get the homeroom name
        homeroom = data.form5.part1.data[i][1];

        // Create a copy of the template with the homeroom name
        fs.copyFile(templateName, `./penyata-akhir-${year}/form5/5${homeroom}-${year}.xlsx`, (err) => {
            if (err) throw err;
            // console.log(`${homeroom}.xlsx was copied to specified directory`);
        });
    }

    console.log('Copied all homerooms from form 5');
}

async function writeDataToCopyForm5() {
    for (let i = 0; i < 18; i++) {
        // Get the homeroom name
        homeroom = data.form5.part1.data[i][1];

        // Open the copy

        const workbook = new ExcelJS.Workbook();

        await workbook.xlsx.readFile(`./penyata-akhir-${year}/form5/5${homeroom}-${year}.xlsx`);

        const worksheet = workbook.getWorksheet('Sheet1');

        // Write the data to the copy

        // Change the homeroom name

        const cell = worksheet.getCell('B8');

        cell.value = `HOMEROOM 5${homeroom} ${year}`;

        // Part 1

        cellpt1_1 = worksheet.getCell('D16');
        cellpt1_1.value = data.form5.part1.data[i][3] || 0;

        cellpt1_2 = worksheet.getCell('E16');
        cellpt1_2.value = data.form5.part1.data[i][4] || 0;

        cellpt1_3 = worksheet.getCell('D17');
        cellpt1_3.value = data.form5.part1.data[i][5] || 0;

        cellpt1_4 = worksheet.getCell('E17');
        cellpt1_4.value = data.form5.part1.data[i][6] || 0;

        cellpt1_5 = worksheet.getCell('D18');
        cellpt1_5.value = data.form5.part1.data[i][7] || 0;

        cellpt1_6 = worksheet.getCell('E18');
        cellpt1_6.value = data.form5.part1.data[i][8] || 0;

        cellpt1_7 = worksheet.getCell('D19');
        cellpt1_7.value = data.form5.part1.data[i][9] || 0;

        cellpt1_8 = worksheet.getCell('E19');
        cellpt1_8.value = data.form5.part1.data[i][10] || 0;

        // Part 2 (Cells D22, E22, D23, E23, D24, E24, D25, E25)

        cellpt2_1 = worksheet.getCell('D22');
        cellpt2_1.value = data.form5.part2.data[i][0] || 0;

        cellpt2_2 = worksheet.getCell('E22');
        cellpt2_2.value = data.form5.part2.data[i][1] || 0;

        cellpt2_3 = worksheet.getCell('D23');
        cellpt2_3.value = data.form5.part2.data[i][2] || 0;

        cellpt2_4 = worksheet.getCell('E23');
        cellpt2_4.value = data.form5.part2.data[i][3] || 0;

        cellpt2_5 = worksheet.getCell('D24');
        cellpt2_5.value = data.form5.part2.data[i][4] || 0;

        cellpt2_6 = worksheet.getCell('E24');
        cellpt2_6.value = data.form5.part2.data[i][5] || 0;

        cellpt2_7 = worksheet.getCell('D25');
        cellpt2_7.value = data.form5.part2.data[i][6] || 0;

        cellpt2_8 = worksheet.getCell('E25');
        cellpt2_8.value = data.form5.part2.data[i][7] || 0;

        // Part 3 (Cells E28, E29, E30, E31, )

        cellpt3_1 = worksheet.getCell('E28');
        cellpt3_1.value = data.form5.part3.data[i][0] + data.form5.part3.data[i][1] + data.form5.part3.data[i][2] || 0;

        cellpt3_2 = worksheet.getCell('E29');
        cellpt3_2.value = data.form5.part3.data[i][3] + data.form5.part3.data[i][4] + data.form5.part3.data[i][5] || 0;

        cellpt3_3 = worksheet.getCell('E30');
        cellpt3_3.value = data.form5.part3.data[i][6] + data.form5.part3.data[i][7] + data.form5.part3.data[i][8] || 0;

        cellpt3_4 = worksheet.getCell('E31');
        cellpt3_4.value = 0

        // Part 4 (Cells D34, E34, D35, E35, D36, E36, D37, E37, D38, E38, D39, E39, D40, E40) depending on how many pertandingans there are

        // Check number of pertandingans and only write to their respective cells, going down for each pertandingan

        cellpt4_1 = worksheet.getCell('D34');
        cellpt4_1.value = data.form5.part4.data[i + 1][0] || 0;

        cellpt4_2 = worksheet.getCell('E34');
        cellpt4_2.value = data.form5.part4.data[i + 1][1] || 0;

        cellpt4_3 = worksheet.getCell('D35');
        cellpt4_3.value = data.form5.part4.data[i + 1][2] || 0;

        cellpt4_4 = worksheet.getCell('E35');
        cellpt4_4.value = data.form5.part4.data[i + 1][3] || 0;

        cellpt4_5 = worksheet.getCell('D36');
        cellpt4_5.value = data.form5.part4.data[i + 1][4] || 0;

        cellpt4_6 = worksheet.getCell('E36');
        cellpt4_6.value = data.form5.part4.data[i + 1][5] || 0;

        cellpt4_7 = worksheet.getCell('D37');
        cellpt4_7.value = data.form5.part4.data[i + 1][6] || 0;

        cellpt4_8 = worksheet.getCell('E37');
        cellpt4_8.value = data.form5.part4.data[i + 1][7] || 0;

        cellpt4_9 = worksheet.getCell('D38');
        cellpt4_9.value = data.form5.part4.data[i + 1][8] || 0;

        cellpt4_10 = worksheet.getCell('E38');
        cellpt4_10.value = data.form5.part4.data[i + 1][9] || 0;

        cellpt4_11 = worksheet.getCell('D39');
        cellpt4_11.value = data.form5.part4.data[i + 1][10] || 0;

        cellpt4_12 = worksheet.getCell('E39');
        cellpt4_12.value = data.form5.part4.data[i + 1][11] || 0;

        cellpt4_13 = worksheet.getCell('D40');
        cellpt4_13.value = data.form5.part4.data[i + 1][12] || 0;

        cellpt4_14 = worksheet.getCell('E40');
        cellpt4_14.value = data.form5.part4.data[i + 1][13] || 0;

        cellpt4_15 = worksheet.getCell('D41');
        cellpt4_15.value = data.form5.part4.data[i + 1][14] || 0;

        cellpt4_16 = worksheet.getCell('E41');
        cellpt4_16.value = data.form5.part4.data[i + 1][15] || 0;

        form5PertandinganNames = data.form5.part4.data[0];

        // Fil in the pertandingan names from C34 until C40

        for (let j = 0; j < form5PertandinganNames.length; j++) {
            cellpt4_9 = worksheet.getCell(`C${34 + j}`);
            cellpt4_9.value = form5PertandinganNames[j];

            cellpt4_10 = worksheet.getCell('C41');
            cellpt4_10.value = '';
        }

        // Write the data to the copy

        await workbook.xlsx.writeFile(`./penyata-akhir-${year}/form5/5${homeroom}-${year}.xlsx`);

    }

    console.log('Done writing data to copies of Form 5')

}

// Run the functions

copyFromTemplateForm1();
writeDataToCopyForm1();

copyFromTemplateForm2();
writeDataToCopyForm2();

copyFromTemplateForm3();
writeDataToCopyForm3();

copyFromTemplateForm4();
writeDataToCopyForm4();

copyFromTemplateForm5();
writeDataToCopyForm5();





