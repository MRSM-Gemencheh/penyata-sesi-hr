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

        cell.value = `HOMEROOM 1${homeroom} 2022`;

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
    
    console.log('Done writing data to copies of Form 1');

    // Delete all of the templates

    for (let i = 0; i < 15; i++) {
        homeroom = data.form1.part1.data[i][1];

        fs.unlink(`./penyata-akhir/form1/1${homeroom}-2022.xlsx`, (err) => {
            if (err) throw err;
            // console.log(`${homeroom}.xlsx was deleted`);
        });
    }
}


// Form 2

async function copyFromTemplateForm2() {
    for (let i = 0; i < 18; i++) {
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

async function writeDataToCopyForm2() {
    for (let i = 0; i < 18; i++) {
        // Get the homeroom name
        homeroom = data.form2.part1.data[i][1];

        // Open the copy

        const workbook = new ExcelJS.Workbook();

        await workbook.xlsx.readFile(`./penyata-akhir/form2/2${homeroom}-2022.xlsx`);

        const worksheet = workbook.getWorksheet('Sheet1');

        // Write the data to the copy

        // Change the homeroom name

        const cell = worksheet.getCell('B8');

        cell.value = `HOMEROOM 2${homeroom} 2022`;

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
        // Kolokium, HB, Scrabble, Docu, Cerpen
        cellpt4_1 = worksheet.getCell('D32');
        cellpt4_1.value = data.form2.part4.data[i][0];

        cellpt4_2 = worksheet.getCell('E32');
        cellpt4_2.value = data.form2.part4.data[i][1];

        cellpt4_3 = worksheet.getCell('D33');
        cellpt4_3.value = data.form2.part4.data[i][2];

        cellpt4_4 = worksheet.getCell('E33');
        cellpt4_4.value = data.form2.part4.data[i][3];

        chgtextcellpt4 = worksheet.getCell('C34');
        chgtextcellpt4.value = 'Scrabble';

        cellpt4_5 = worksheet.getCell('D34');
        cellpt4_5.value = data.form2.part4.data[i][4];

        cellpt4_6 = worksheet.getCell('E34');
        cellpt4_6.value = data.form2.part4.data[i][5];

        cellpt4_7 = worksheet.getCell('D36');
        cellpt4_7.value = data.form2.part4.data[i][6];

        cellpt4_8 = worksheet.getCell('E36');
        cellpt4_8.value = data.form2.part4.data[i][7];

        cellpt4_9 = worksheet.getCell('D37');
        cellpt4_9.value = data.form2.part4.data[i][8];



        // Write the data to the copy

        await workbook.xlsx.writeFile(`./penyata-akhir/form2/penyata-akhir-2${homeroom}-2022.xlsx`);

        

    }

    console.log('Done writing data to copies of Form 2')

    // Delete all the copies of the template

    for (let i = 0; i < 18; i++) {
        homeroom = data.form2.part1.data[i][1];

        fs.unlink(`./penyata-akhir/form2/2${homeroom}-2022.xlsx`, (err) => {
            if (err) throw err;
            // console.log(`Deleted 2${homeroom}-2022.xlsx`);
        });
    }


}

// Form 3

async function copyFromTemplateForm3() {
    for (let i = 0; i < 18; i++) {
        // Get the homeroom name
        homeroom = data.form3.part1.data[i][1];
    
        // Create a copy of the template with the homeroom name
        fs.copyFile(templateName, `./penyata-akhir/form3/3${homeroom}-2022.xlsx`, (err) => {
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

        await workbook.xlsx.readFile(`./penyata-akhir/form3/3${homeroom}-2022.xlsx`);

        const worksheet = workbook.getWorksheet('Sheet1');

        // Write the data to the copy

        // Change the homeroom name

        const cell = worksheet.getCell('B8');

        cell.value = `HOMEROOM 3${homeroom} 2022`;

        // Part 1

        cellpt1_1 = worksheet.getCell('D16');
        cellpt1_1.value = data.form3.part1.data[i][3];

        cellpt1_2 = worksheet.getCell('E16');
        cellpt1_2.value = data.form3.part1.data[i][4];

        cellpt1_3 = worksheet.getCell('D17');
        cellpt1_3.value = data.form3.part1.data[i][5];

        cellpt1_4 = worksheet.getCell('E17');
        cellpt1_4.value = data.form3.part1.data[i][6];

        cellpt1_5 = worksheet.getCell('D18');
        cellpt1_5.value = data.form3.part1.data[i][7] + data.form3.part1.data[i][9];

        cellpt1_6 = worksheet.getCell('E18');
        cellpt1_6.value = data.form3.part1.data[i][8] + data.form3.part1.data[i][10];
    
        // Part 2

        cellpt2_1 = worksheet.getCell('E21');
        cellpt2_1.value = data.form3.part3.data[i][0] + data.form3.part3.data[i][1] + data.form3.part3.data[i][2];

        cellpt2_2 = worksheet.getCell('E22');
        cellpt2_2.value = data.form3.part3.data[i][3] + data.form3.part3.data[i][4] + data.form3.part3.data[i][5];

        cellpt2_3 = worksheet.getCell('E23');
        cellpt2_3.value = data.form3.part3.data[i][6] + data.form3.part3.data[i][7] + data.form3.part3.data[i][8];

        // Part 3

        cellpt3_1 = worksheet.getCell('D26');
        cellpt3_1.value = data.form3.part2.data[i][0];

        cellpt3_2 = worksheet.getCell('E26');
        cellpt3_2.value = data.form3.part2.data[i][1];

        cellpt3_3 = worksheet.getCell('D27');
        cellpt3_3.value = data.form3.part2.data[i][2];

        cellpt3_4 = worksheet.getCell('E27');
        cellpt3_4.value = data.form3.part2.data[i][3];

        cellpt3_5 = worksheet.getCell('D28');
        cellpt3_5.value = data.form3.part2.data[i][4];

        cellpt3_6 = worksheet.getCell('E28');
        cellpt3_6.value = data.form3.part2.data[i][5];

        cellpt3_7 = worksheet.getCell('D29');
        cellpt3_7.value = data.form3.part2.data[i][6];

        cellpt3_8 = worksheet.getCell('E29');
        cellpt3_8.value = data.form3.part2.data[i][7];

        // Part 4
        // Kolokium, HB, Scrabble, Docu, Cerpen
        cellpt4_1 = worksheet.getCell('D32');
        cellpt4_1.value = data.form3.part4.data[i][0];

        cellpt4_2 = worksheet.getCell('E32');
        cellpt4_2.value = data.form3.part4.data[i][1];

        cellpt4_3 = worksheet.getCell('D33');
        cellpt4_3.value = data.form3.part4.data[i][2];

        cellpt4_4 = worksheet.getCell('E33');
        cellpt4_4.value = data.form3.part4.data[i][3];

        chgtextcellpt4 = worksheet.getCell('C34');
        chgtextcellpt4.value = 'Scrabble';

        cellpt4_5 = worksheet.getCell('D34');
        cellpt4_5.value = data.form3.part4.data[i][4];

        cellpt4_6 = worksheet.getCell('E34');
        cellpt4_6.value = data.form3.part4.data[i][5];

        cellpt4_7 = worksheet.getCell('D36');
        cellpt4_7.value = data.form3.part4.data[i][6];

        cellpt4_8 = worksheet.getCell('E36');
        cellpt4_8.value = data.form3.part4.data[i][7];

        cellpt4_9 = worksheet.getCell('D37');
        cellpt4_9.value = data.form3.part4.data[i][8];



        // Write the data to the copy

        await workbook.xlsx.writeFile(`./penyata-akhir/form3/penyata-akhir-3${homeroom}-2022.xlsx`);

    }

    console.log('Done writing data to copies of Form 3')

    // Delete all templates 

    for (let i = 0; i < 18; i++) {
        homeroom = data.form3.part1.data[i][1];
        fs.unlink(`./penyata-akhir/form3/3${homeroom}-2022.xlsx`, (err) => {
            if (err) throw err;
            // console.log(`Deleted 3${homeroom}-2022.xlsx`);
        });
    }

}

// Form 4

async function copyFromTemplateForm4() {
    for (let i = 0; i < 15; i++) {
        // Get the homeroom name
        homeroom = data.form4.part1.data[i][1];
    
        // Create a copy of the template with the homeroom name
        fs.copyFile(templateName, `./penyata-akhir/form4/4${homeroom}-2022.xlsx`, (err) => {
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

        await workbook.xlsx.readFile(`./penyata-akhir/form4/4${homeroom}-2022.xlsx`);

        const worksheet = workbook.getWorksheet('Sheet1');

        // Write the data to the copy

        // Change the homeroom name

        const cell = worksheet.getCell('B8');

        cell.value = `HOMEROOM 4${homeroom} 2022`;

        // Part 1

        cellpt1_1 = worksheet.getCell('D16');
        cellpt1_1.value = data.form4.part1.data[i][3];

        cellpt1_2 = worksheet.getCell('E16');
        cellpt1_2.value = data.form4.part1.data[i][4];

        cellpt1_3 = worksheet.getCell('D17');
        cellpt1_3.value = data.form4.part1.data[i][5];

        cellpt1_4 = worksheet.getCell('E17');
        cellpt1_4.value = data.form4.part1.data[i][6];

        cellpt1_5 = worksheet.getCell('D18');
        cellpt1_5.value = data.form4.part1.data[i][7] + data.form4.part1.data[i][9];

        cellpt1_6 = worksheet.getCell('E18');
        cellpt1_6.value = data.form4.part1.data[i][8] + data.form4.part1.data[i][10];
    
        // Part 2

        cellpt2_1 = worksheet.getCell('E21');
        cellpt2_1.value = data.form4.part3.data[i][0] + data.form4.part3.data[i][1] + data.form4.part3.data[i][2];

        cellpt2_2 = worksheet.getCell('E22');
        cellpt2_2.value = data.form4.part3.data[i][3] + data.form4.part3.data[i][4] + data.form4.part3.data[i][5];

        cellpt2_3 = worksheet.getCell('E23');
        cellpt2_3.value = data.form4.part3.data[i][6] + data.form4.part3.data[i][7] + data.form4.part3.data[i][8];

        // Part 3

        cellpt3_1 = worksheet.getCell('D26');
        cellpt3_1.value = data.form4.part2.data[i][0];

        cellpt3_2 = worksheet.getCell('E26');
        cellpt3_2.value = data.form4.part2.data[i][1];

        cellpt3_3 = worksheet.getCell('D27');
        cellpt3_3.value = data.form4.part2.data[i][2];

        cellpt3_4 = worksheet.getCell('E27');
        cellpt3_4.value = data.form4.part2.data[i][3];

        cellpt3_5 = worksheet.getCell('D28');
        cellpt3_5.value = data.form4.part2.data[i][4];

        cellpt3_6 = worksheet.getCell('E28');
        cellpt3_6.value = data.form4.part2.data[i][5];

        cellpt3_7 = worksheet.getCell('D29');
        cellpt3_7.value = data.form4.part2.data[i][6];

        cellpt3_8 = worksheet.getCell('E29');
        cellpt3_8.value = data.form4.part2.data[i][7];

        // Part 4
        // Kolokium, HB, Scrabble, Docu, Cerpen
        cellpt4_1 = worksheet.getCell('D32');
        cellpt4_1.value = data.form4.part4.data[i][0];

        cellpt4_2 = worksheet.getCell('E32');
        cellpt4_2.value = data.form4.part4.data[i][1];

        cellpt4_3 = worksheet.getCell('D33');
        cellpt4_3.value = data.form4.part4.data[i][2];

        cellpt4_4 = worksheet.getCell('E33');
        cellpt4_4.value = data.form4.part4.data[i][3];

        cellpt4_7 = worksheet.getCell('D36');
        cellpt4_7.value = data.form4.part4.data[i][4];

        cellpt4_8 = worksheet.getCell('E36');
        cellpt4_8.value = data.form4.part4.data[i][5];

        cellpt4_9 = worksheet.getCell('D37');
        cellpt4_9.value = data.form4.part4.data[i][6];

        // Write the data to the copy

        await workbook.xlsx.writeFile(`./penyata-akhir/form4/penyata-akhir-4${homeroom}-2022.xlsx`);

    }

    console.log('Done writing data to copies of Form 4')

    // Delete all templates 

    for (let i = 0; i < 15; i++) {
        homeroom = data.form4.part1.data[i][1];
        fs.unlink(`./penyata-akhir/form4/4${homeroom}-2022.xlsx`, (err) => {
            if (err) throw err;
            // console.log(`Deleted 4${homeroom}-2022.xlsx`);
        });
    }

}


// Form 5

async function copyFromTemplateForm5() {
    for (let i = 0; i < 18; i++) {
        // Get the homeroom name
        homeroom = data.form5.part1.data[i][1];
    
        // Create a copy of the template with the homeroom name
        fs.copyFile(templateName, `./penyata-akhir/form5/5${homeroom}-2022.xlsx`, (err) => {
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

        await workbook.xlsx.readFile(`./penyata-akhir/form5/5${homeroom}-2022.xlsx`);

        const worksheet = workbook.getWorksheet('Sheet1');

        // Write the data to the copy

        // Change the homeroom name

        const cell = worksheet.getCell('B8');

        cell.value = `HOMEROOM 5${homeroom} 2022`;

        // Part 1

        cellpt1_1 = worksheet.getCell('D16');
        cellpt1_1.value = data.form5.part1.data[i][3];

        cellpt1_2 = worksheet.getCell('E16');
        cellpt1_2.value = data.form5.part1.data[i][4];

        cellpt1_3 = worksheet.getCell('D17');
        cellpt1_3.value = data.form5.part1.data[i][5];

        cellpt1_4 = worksheet.getCell('E17');
        cellpt1_4.value = data.form5.part1.data[i][6];

        cellpt1_5 = worksheet.getCell('D18');
        cellpt1_5.value = data.form5.part1.data[i][7] + data.form5.part1.data[i][9];

        cellpt1_6 = worksheet.getCell('E18');
        cellpt1_6.value = data.form5.part1.data[i][8] + data.form5.part1.data[i][10];
    
        // Part 2

        cellpt2_1 = worksheet.getCell('E21');
        cellpt2_1.value = data.form5.part3.data[i][0] + data.form5.part3.data[i][1] + data.form5.part3.data[i][2];

        cellpt2_2 = worksheet.getCell('E22');
        cellpt2_2.value = data.form5.part3.data[i][3] + data.form5.part3.data[i][4] + data.form5.part3.data[i][5];

        cellpt2_3 = worksheet.getCell('E23');
        cellpt2_3.value = data.form5.part3.data[i][6] + data.form5.part3.data[i][7] + data.form5.part3.data[i][8];

        // Part 3

        cellpt3_1 = worksheet.getCell('D26');
        cellpt3_1.value = data.form5.part2.data[i][0];

        cellpt3_2 = worksheet.getCell('E26');
        cellpt3_2.value = data.form5.part2.data[i][1];

        cellpt3_3 = worksheet.getCell('D27');
        cellpt3_3.value = data.form5.part2.data[i][2];

        cellpt3_4 = worksheet.getCell('E27');
        cellpt3_4.value = data.form5.part2.data[i][3];

        cellpt3_5 = worksheet.getCell('D28');
        cellpt3_5.value = data.form5.part2.data[i][4];

        cellpt3_6 = worksheet.getCell('E28');
        cellpt3_6.value = data.form5.part2.data[i][5];

        cellpt3_7 = worksheet.getCell('D29');
        cellpt3_7.value = data.form5.part2.data[i][6];

        cellpt3_8 = worksheet.getCell('E29');
        cellpt3_8.value = data.form5.part2.data[i][7];

        // Part 4
        // Kolokium, HB, Scrabble, Docu, Cerpen
        cellpt4_1 = worksheet.getCell('D32');
        cellpt4_1.value = data.form5.part4.data[i][0];

        cellpt4_2 = worksheet.getCell('E32');
        cellpt4_2.value = data.form5.part4.data[i][1];

        cellpt4_3 = worksheet.getCell('D33');
        cellpt4_3.value = data.form5.part4.data[i][2] + data.form5.part4.data[i][4];

        cellpt4_4 = worksheet.getCell('E33');
        cellpt4_4.value = data.form5.part4.data[i][3] + data.form5.part4.data[i][5];

        cellpt4_5 = worksheet.getCell('D36');
        cellpt4_5.value = data.form5.part4.data[i][6];

        cellpt4_6 = worksheet.getCell('E36');
        cellpt4_6.value = data.form5.part4.data[i][7];

        cellpt4_7 = worksheet.getCell('D37');
        cellpt4_7.value = data.form5.part4.data[i][8];

        // Write the data to the copy

        await workbook.xlsx.writeFile(`./penyata-akhir/form5/penyata-akhir-5${homeroom}-2022.xlsx`);

    }

    console.log('Done writing data to copies of Form 5')

    // Delete all templates 

    for (let i = 0; i < 18; i++) {
        homeroom = data.form5.part1.data[i][1];
        fs.unlink(`./penyata-akhir/form5/5${homeroom}-2022.xlsx`, (err) => {
            if (err) throw err;
            // console.log(`Deleted 4${homeroom}-2022.xlsx`);
        });
    }

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





