const ExcelJS = require('exceljs');
const fs = require('fs');

try {
    if (ExcelJS) {
        console.info("ExcelJS loaded successfully!")
    }
} catch (err) {
    console.error("ExcelJS failed to load")
}

const workbook = new ExcelJS.Workbook();

const dataFileSelector = document.getElementById('data-file-selector');
dataFileSelector.addEventListener('change', (event) => {
    const fileList = event.target.files;
    console.log(fileList);

    const file = fileList[0];
    readDataFile(file);
});


function readDataFile(file) {
    const reader = new FileReader();
    reader.addEventListener('load', (event) => {
        const result = event.target.result;
        readExcelFileFromSystem(result);

        console.log(result);
    });

    reader.addEventListener('progress', (event) => {
        if (event.loaded && event.total) {
            const percent = (event.loaded / event.total) * 100;
            console.log(`Progress: ${Math.round(percent)}`);
        }
    });
    reader.readAsDataURL(file);
}

async function readExcelFileFromSystem(data) {

    await workbook.xlsx.load(data);

    readDataFromFile()

}

async function readDataFromFile() {
    form1Worksheet = workbook.getWorksheet('TING.1')
    form2Worksheet = workbook.getWorksheet('TING.2')
    form3Worksheet = workbook.getWorksheet('TING.3')
    form4Worksheet = workbook.getWorksheet('TING.4')
    form5Worksheet = workbook.getWorksheet('TING.5')

    // Logging all of the actualRowCounts of every form in one line

    console.log("Actual row counts of every form:")
    console.log("Form 1: " + form1Worksheet.actualRowCount + " rows")
    console.log("Form 2: " + form2Worksheet.actualRowCount + " rows")
    console.log("Form 3: " + form3Worksheet.actualRowCount + " rows")
    console.log("Form 4: " + form4Worksheet.actualRowCount + " rows")
    console.log("Form 5: " + form5Worksheet.actualRowCount + " rows")

    // Logging all of the actualColumnCounts of every form
    // If the actualColumntCount of every form is the same, then don't log the column values

    if (form1Worksheet.actualColumnCount == form2Worksheet.actualColumnCount && form1Worksheet.actualColumnCount == form3Worksheet.actualColumnCount && form1Worksheet.actualColumnCount == form4Worksheet.actualColumnCount && form1Worksheet.actualColumnCount == form5Worksheet.actualColumnCount) {

        console.log("All of the actual column counts from all forms are the same!")

    } else {
        console.log("Actual column counts of every form:")
        console.log("Form 1: " + form1Worksheet.actualColumnCount)
        console.log("Form 2: " + form2Worksheet.actualColumnCount)
        console.log("Form 3: " + form3Worksheet.actualColumnCount)
        console.log("Form 4: " + form4Worksheet.actualColumnCount)
        console.log("Form 5: " + form5Worksheet.actualColumnCount)
    }

    // This will affect what rows we'll need to get data from below

    // Part 1

    // Form 1

    form1Part1Data = []


    for (let i = 11; i <= 28; i++) {
        form1Part1Data.push(form1Worksheet.getRow(i).values)
    }


    // Form 2

    form2Part1Data = []


    for (let i = 11; i <= 28; i++) {
        form2Part1Data.push(form2Worksheet.getRow(i).values)
    }


    // Form 3

    form3Part1Data = []


    for (let i = 11; i <= 28; i++) {
        form3Part1Data.push(form3Worksheet.getRow(i).values)
    }

    // Form 4

    form4Part1Data = []


    for (let i = 11; i <= 28; i++) {
        form4Part1Data.push(form4Worksheet.getRow(i).values)
    }

    // Form 5

    form5Part1Data = []


    for (let i = 11; i <= 28; i++) {
        form5Part1Data.push(form5Worksheet.getRow(i).values)
    }


    // Logging the data

    // console.log("Form 1 Part 1 Data:")
    // console.log(form1Part1Data)
    // console.log("Form 2 Part 1 Data:")
    // console.log(form2Part1Data)
    // console.log("Form 3 Part 1 Data:")
    // console.log(form3Part1Data)
    // console.log("Form 4 Part 1 Data:")
    // console.log(form4Part1Data)
    // console.log("Form 5 Part 1 Data:")
    // console.log(form5Part1Data)

    // Part 2

    // Store the data in an array of arrays

    // Form 1

    form1Part2Data = []


    for (let i = 36; i <= 53; i++) {
        form1Part2Data.push(form1Worksheet.getRow(i).values.slice(3, 12))
    }


    // Form 2

    form2Part2Data = []


    for (let i = 36; i <= 53; i++) {
        form2Part2Data.push(form2Worksheet.getRow(i).values.slice(3, 12))
    }


    // Form 3

    form3Part2Data = []


    for (let i = 36; i <= 53; i++) {
        form3Part2Data.push(form3Worksheet.getRow(i).values.slice(3, 12))
    }

    // Form 4

    form4Part2Data = []


    for (let i = 36; i <= 53; i++) {
        form4Part2Data.push(form4Worksheet.getRow(i).values.slice(3, 12))
    }


    // Form 5

    form5Part2Data = []


    for (let i = 36; i <= 53; i++) {
        form5Part2Data.push(form5Worksheet.getRow(i).values.slice(3, 12))
    }


    // Logging the data

    // console.log("Form 1 Part 2 Data:")
    // console.log(form1Part2Data)
    // console.log("Form 2 Part 2 Data:")
    // console.log(form2Part2Data)
    // console.log("Form 3 Part 2 Data:")
    // console.log(form3Part2Data)
    // console.log("Form 4 Part 2 Data:")
    // console.log(form4Part2Data)
    // console.log("Form 5 Part 2 Data:")
    // console.log(form5Part2Data)

    // Part 3

    // Store the data in an array of arrays

    // Form 1

    form1Part3Data = []


    for (let i = 61; i <= 78; i++) {
        form1Part3Data.push(form1Worksheet.getRow(i).values.slice(3, 13))
    }



    // Form 2

    form2Part3Data = []


    for (let i = 61; i <= 78; i++) {
        form2Part3Data.push(form2Worksheet.getRow(i).values.slice(3, 13))
    }

    // Form 3

    form3Part3Data = []


    for (let i = 61; i <= 78; i++) {
        form3Part3Data.push(form3Worksheet.getRow(i).values.slice(3, 13))
    }


    // Form 4

    form4Part3Data = []


    for (let i = 61; i <= 78; i++) {
        form4Part3Data.push(form4Worksheet.getRow(i).values.slice(3, 13))
    }


    // Form 5

    form5Part3Data = []

    for (let i = 61; i <= 78; i++) {
        form5Part3Data.push(form5Worksheet.getRow(i).values.slice(3, 13))
    }

    // Logging the data

    // console.log("Form 1 Part 3 Data:")
    // console.log(form1Part3Data)
    // console.log("Form 2 Part 3 Data:")
    // console.log(form2Part3Data)
    // console.log("Form 3 Part 3 Data:")
    // console.log(form3Part3Data)
    // console.log("Form 4 Part 3 Data:")
    // console.log(form4Part3Data)
    // console.log("Form 5 Part 3 Data:")
    // console.log(form5Part3Data)

    // Part 4

    // Store the data in an array of arrays

    // Form 1

    form1Part4Data = []

    // Get all of the pertandingan names from row 84, not pushing into the form1Part4Data array if duplicate

    let form1PertandinganNames = []

    for (let i = 3; i <= 19; i++) {
        form1PertandinganNames.push(form1Worksheet.getRow(84).values[i])
    }

    // Remove duplicates from the array

    form1PertandinganNames = [...new Set(form1PertandinganNames)]

    // Also remove instances of 'NAMA PERTANDINGAN' and 'JUMLAH' from the array, replace them with null

    form1PertandinganNames = form1PertandinganNames.map((item) => {
        if (item === 'NAMA PERTANDINGAN' || item === 'JUMLAH' || item === 'Jumlah' || item == 'Nama Pertandingan') {

            return null
        } else {
            return item.toLowerCase(); // Change all characters to lowercase
        }
    })

    // Also change every item to use Title Case

    form1PertandinganNames = form1PertandinganNames.map((item) => {
        if (item !== null) {
            return item.replace(/\w\S*/g, (w) => (w.replace(/^\w/, (c) => c.toUpperCase())))
        } else {
            return null
        }
    })

    form1Part4Data.push(form1PertandinganNames)

    for (let i = 86; i <= 103; i++) {
        form1Part4Data.push(form1Worksheet.getRow(i).values.slice(3, 19))
    }


    // Form 2

    form2Part4Data = []

    // Get all of the pertandingan names from row 84, not pushing into the form2Part4Data array if duplicate

    let form2PertandinganNames = []

    for (let i = 3; i <= 19; i++) {
        form2PertandinganNames.push(form2Worksheet.getRow(84).values[i])
    }

    // Remove duplicates from the array

    form2PertandinganNames = [...new Set(form2PertandinganNames)]

    // Also remove instances of 'NAMA PERTANDINGAN' and 'JUMLAH' from the array, replace them with null

    form2PertandinganNames = form2PertandinganNames.map((item) => {
        if (item === 'NAMA PERTANDINGAN' || item === 'JUMLAH' || item === 'Jumlah' || item == 'Nama Pertandingan') {
            return null
        } else {
            return item.toLowerCase(); // Change all characters to lowercase
        }
    })

    // Also change every item to use Title Case

    form2PertandinganNames = form2PertandinganNames.map((item) => {
        if (item !== null) {
            return item.replace(/\w\S*/g, (w) => (w.replace(/^\w/, (c) => c.toUpperCase())))
        } else {
            return null
        }
    })

    form2Part4Data.push(form2PertandinganNames)

    for (let i = 86; i <= 103; i++) {
        form2Part4Data.push(form2Worksheet.getRow(i).values.slice(3, 19))
    }


    // Form 3

    form3Part4Data = []

    // Get all of the pertandingan names from row 84, not pushing into the form3Part4Data array if duplicate

    let form3PertandinganNames = []

    for (let i = 3; i <= 19; i++) {
        form3PertandinganNames.push(form3Worksheet.getRow(84).values[i])
    }

    // Remove duplicates from the array

    form3PertandinganNames = [...new Set(form3PertandinganNames)]

    // Also remove instances of 'NAMA PERTANDINGAN' and 'JUMLAH' from the array, replace them with null

    form3PertandinganNames = form3PertandinganNames.map((item) => {
        if (item === 'NAMA PERTANDINGAN' || item === 'JUMLAH' || item === 'Jumlah' || item == 'Nama Pertandingan') {
            return null
        } else {
            return item.toLowerCase(); // Change all characters to lowercase
        }
    })

    // Also change every item to use Title Case

    form3PertandinganNames = form3PertandinganNames.map((item) => {
        if (item !== null) {
            return item.replace(/\w\S*/g, (w) => (w.replace(/^\w/, (c) => c.toUpperCase())))
        } else {
            return null
        }
    })

    form3Part4Data.push(form3PertandinganNames)

    for (let i = 86; i <= 103; i++) {
        form3Part4Data.push(form3Worksheet.getRow(i).values.slice(3, 19))
    }


    // Form 4

    form4Part4Data = []

    // Get all of the pertandingan names from row 84, not pushing into the form4Part4Data array if duplicate

    let form4PertandinganNames = []

    for (let i = 3; i <= 19; i++) {
        form4PertandinganNames.push(form4Worksheet.getRow(84).values[i])
    }

    // Remove duplicates from the array

    form4PertandinganNames = [...new Set(form4PertandinganNames)]

    // Also remove instances of 'NAMA PERTANDINGAN' and 'JUMLAH' from the array, replace them with null

    form4PertandinganNames = form4PertandinganNames.map((item) => {
        if (item === 'NAMA PERTANDINGAN' || item === 'JUMLAH' || item === 'Jumlah' || item == 'Nama Pertandingan') {
            return null
        } else {
            return item.toLowerCase(); // Change all characters to lowercase
        }
    })

    // Also change every item to use Title Case

    form4PertandinganNames = form4PertandinganNames.map((item) => {
        if (item !== null) {
            return item.replace(/\w\S*/g, (w) => (w.replace(/^\w/, (c) => c.toUpperCase())))
        } else {
            return null
        }
    })

    form4Part4Data.push(form4PertandinganNames)

    for (let i = 86; i <= 103; i++) {
        form4Part4Data.push(form4Worksheet.getRow(i).values.slice(3, 19))
    }


    // Form 5

    form5Part4Data = []

    // Get all of the pertandingan names from row 84, not pushing into the form5Part4Data array if duplicate

    let form5PertandinganNames = []

    for (let i = 3; i <= 19; i++) {
        form5PertandinganNames.push(form5Worksheet.getRow(84).values[i])
    }

    // Remove duplicates from the array

    form5PertandinganNames = [...new Set(form5PertandinganNames)]

    // Also remove instances of 'NAMA PERTANDINGAN' and 'JUMLAH' from the array, replace them with null

    form5PertandinganNames = form5PertandinganNames.map((item) => {
        if (item === 'NAMA PERTANDINGAN' || item === 'JUMLAH' || item === 'Jumlah' || item == 'Nama Pertandingan') {
            return null
        } else {
            return item.toLowerCase(); // Change all characters to lowercase
        }
    })

    // Also change every item to use Title Case

    form5PertandinganNames = form5PertandinganNames.map((item) => {
        if (item !== null) {
            return item.replace(/\w\S*/g, (w) => (w.replace(/^\w/, (c) => c.toUpperCase())))
        } else {
            return null
        }
    })

    form5Part4Data.push(form5PertandinganNames)

    for (let i = 86; i <= 103; i++) {
        form5Part4Data.push(form5Worksheet.getRow(i).values.slice(3, 19))
    }


    // Logging the data

    // console.log("Form 1 Part 4 Data:")
    // console.log(form1Part4Data)
    // console.log("Form 2 Part 4 Data:")
    // console.log(form2Part4Data)
    // console.log("Form 3 Part 4 Data:")
    // console.log(form3Part4Data)
    // console.log("Form 4 Part 4 Data:")
    // console.log(form4Part4Data)
    // console.log("Form 5 Part 4 Data:")
    // console.log(form5Part4Data)

    await exportToJSON()
}


// Create a JSON object

const data = {
    "form1": {
        "part1": {
            "data": form1Part1Data
        },
        "part2": {
            "data": form1Part2Data
        },
        "part3": {
            "data": form1Part3Data
        },
        "part4": {
            "data": form1Part4Data
        }

    },
    "form2": {
        "part1": {
            "data": form2Part1Data
        },
        "part2": {
            "data": form2Part2Data
        },
        "part3": {
            "data": form2Part3Data
        },
        "part4": {
            "data": form2Part4Data
        }

    },
    "form3": {
        "part1": {
            "data": form3Part1Data
        },
        "part2": {
            "data": form3Part2Data
        },
        "part3": {
            "data": form3Part3Data
        },
        "part4": {
            "data": form3Part4Data
        }

    },
    "form4": {
        "part1": {
            "data": form4Part1Data
        },
        "part2": {
            "data": form4Part2Data
        },
        "part3": {
            "data": form4Part3Data
        },
        "part4": {
            "data": form4Part4Data
        }

    },
    "form5": {
        "part1": {
            "data": form5Part1Data
        },
        "part2": {
            "data": form5Part2Data
        },
        "part3": {
            "data": form5Part3Data
        },
        "part4": {
            "data": form5Part4Data
        }

    }
}


const templateFileSelector = document.getElementById('template-file-selector');
templateFileSelector.addEventListener('change', (event) => {
    const fileList = event.target.files;
    console.log(fileList);

    const file = fileList[0];
    readDataFile(file);
});


function readDataFile(file) {
    const reader = new FileReader();
    reader.addEventListener('load', (event) => {
        const result = event.target.result;

        console.log(result);
    });

    reader.addEventListener('progress', (event) => {
        if (event.loaded && event.total) {
            const percent = (event.loaded / event.total) * 100;
            console.log(`Progress: ${Math.round(percent)}`);
        }
    });
    reader.readAsDataURL(file);
}

// CODE BREAKS HERE


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