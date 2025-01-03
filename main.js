import ExcelJS from 'exceljs';
const JSZip = require('jszip');
const { saveAs } = require('file-saver');

try {
    if (ExcelJS) {
        console.info("ExcelJS loaded successfully!")
    }
} catch (err) {
    console.error("ExcelJS failed to load, try reloading?")
}

// Add event listener for data file input
const dataFileInput = document.getElementById('data-file-selector');
dataFileInput.addEventListener('change', handleDataFileUpload);

// Add event listener for template file input
const templateFileInput = document.getElementById('template-file-selector');
templateFileInput.addEventListener('change', handleTemplateFileUpload);

// Function to handle data file upload
function handleDataFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        processExcelData(data);
    };
    reader.readAsArrayBuffer(file);
}

// Function to handle template file upload
function handleTemplateFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        processExcelTemplate(data);
    };
    reader.readAsArrayBuffer(file);
}

const year = new Date().getFullYear(); // To include in the final zip file name
let meritData = {} // This will be populated with data from the excel file soon
let statementFiles = [] // Generated excel files will be pushed to this array

// Function to process data from data file
// This extracts data from the file and pushes the data into the `meritData` JSON object
function processExcelData(data) {
    const workbook = new ExcelJS.Workbook();
    workbook.xlsx.load(data).then(workbook => {

        console.log("Data file loaded successfully:", workbook);

        // TODO: Check if the uploaded excel file is correct
        // This can be done by checking certain columns if they exist or not

        const form1Worksheet = workbook.getWorksheet('TING.1')
        const form2Worksheet = workbook.getWorksheet('TING.2')
        const form3Worksheet = workbook.getWorksheet('TING.3')
        const form4Worksheet = workbook.getWorksheet('TING.4')
        const form5Worksheet = workbook.getWorksheet('TING.5')

        // Part 1

        let form1Part1Data, form2Part1Data, form3Part1Data, form4Part1Data, form5Part1Data
        form1Part1Data = []
        form2Part1Data = []
        form3Part1Data = []
        form4Part1Data = []
        form5Part1Data = []

        for (let i = 11; i <= 28; i++) {
            form1Part1Data.push(form1Worksheet.getRow(i).values)
            form2Part1Data.push(form2Worksheet.getRow(i).values)
            form3Part1Data.push(form3Worksheet.getRow(i).values)
            form4Part1Data.push(form4Worksheet.getRow(i).values)
            form5Part1Data.push(form5Worksheet.getRow(i).values)
        }

        // Part 2

        let form1Part2Data, form2Part2Data, form3Part2Data, form4Part2Data, form5Part2Data
        form1Part2Data = []
        form2Part2Data = []
        form3Part2Data = []
        form4Part2Data = []
        form5Part2Data = []

        for (let i = 36; i <= 53; i++) {
            form1Part2Data.push(form1Worksheet.getRow(i).values.slice(3, 12))
            form2Part2Data.push(form2Worksheet.getRow(i).values.slice(3, 12))
            form3Part2Data.push(form3Worksheet.getRow(i).values.slice(3, 12))
            form4Part2Data.push(form4Worksheet.getRow(i).values.slice(3, 12))
            form5Part2Data.push(form5Worksheet.getRow(i).values.slice(3, 12))
        }

        // Part 3

        let form1Part3Data, form2Part3Data, form3Part3Data, form4Part3Data, form5Part3Data
        form1Part3Data = []
        form2Part3Data = []
        form3Part3Data = []
        form4Part3Data = []
        form5Part3Data = []

        for (let i = 61; i <= 78; i++) {
            form1Part3Data.push(form1Worksheet.getRow(i).values.slice(3, 13))
            form2Part3Data.push(form2Worksheet.getRow(i).values.slice(3, 13))
            form3Part3Data.push(form3Worksheet.getRow(i).values.slice(3, 13))
            form4Part3Data.push(form4Worksheet.getRow(i).values.slice(3, 13))
            form5Part3Data.push(form5Worksheet.getRow(i).values.slice(3, 13))
        }

        // Part 4

        // Form 1

        let form1PertandinganNames, form2PertandinganNames, form3PertandinganNames, form4PertandinganNames, form5PertandinganNames
        form1PertandinganNames = []
        form2PertandinganNames = []
        form3PertandinganNames = []
        form4PertandinganNames = []
        form5PertandinganNames = []

        for (let i = 3; i <= 19; i++) {
            // Get all of the pertandingan names from row 84
            form1PertandinganNames.push(form1Worksheet.getRow(84).values[i])
            form2PertandinganNames.push(form2Worksheet.getRow(84).values[i])
            form3PertandinganNames.push(form3Worksheet.getRow(84).values[i])
            form4PertandinganNames.push(form4Worksheet.getRow(84).values[i])
            form5PertandinganNames.push(form5Worksheet.getRow(84).values[i])
        }

        // Remove duplicates from the array
        form1PertandinganNames = [...new Set(form1PertandinganNames)]
        form2PertandinganNames = [...new Set(form2PertandinganNames)]
        form3PertandinganNames = [...new Set(form3PertandinganNames)]
        form4PertandinganNames = [...new Set(form4PertandinganNames)]
        form5PertandinganNames = [...new Set(form5PertandinganNames)]

        // Remove instances of 'NAMA PERTANDINGAN' and 'JUMLAH' from the array, replace them with null
        // Also change every valid item to use Title Case

        function cleanNamesAndTitleCase(array) {
            array = array.map((item) => {
                if (item === 'NAMA PERTANDINGAN' || item === 'JUMLAH' || item === 'Jumlah' || item == 'Nama Pertandingan') {
                    return null
                } else {
                    // Lowercase first
                    item = item.toLowerCase()
                    // Capitalize first letter of every word
                    item = item.replace(/\b\w/g, l => l.toUpperCase())
                    return item
                }
            })

            return array
        }

        form1PertandinganNames = cleanNamesAndTitleCase(form1PertandinganNames)
        form2PertandinganNames = cleanNamesAndTitleCase(form2PertandinganNames)
        form3PertandinganNames = cleanNamesAndTitleCase(form3PertandinganNames)
        form4PertandinganNames = cleanNamesAndTitleCase(form4PertandinganNames)
        form5PertandinganNames = cleanNamesAndTitleCase(form5PertandinganNames)


        let form1Part4Data, form2Part4Data, form3Part4Data, form4Part4Data, form5Part4Data
        form1Part4Data = []
        form2Part4Data = []
        form3Part4Data = []
        form4Part4Data = []
        form5Part4Data = []

        form1Part4Data.push(form1PertandinganNames)
        form2Part4Data.push(form2PertandinganNames)
        form3Part4Data.push(form3PertandinganNames)
        form4Part4Data.push(form4PertandinganNames)
        form5Part4Data.push(form5PertandinganNames)

        for (let i = 86; i <= 103; i++) {
            form1Part4Data.push(form1Worksheet.getRow(i).values.slice(3, 19))
            form2Part4Data.push(form2Worksheet.getRow(i).values.slice(3, 19))
            form3Part4Data.push(form3Worksheet.getRow(i).values.slice(3, 19))
            form4Part4Data.push(form4Worksheet.getRow(i).values.slice(3, 19))
            form5Part4Data.push(form5Worksheet.getRow(i).values.slice(3, 19))
        }

        meritData = {
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
    }

    ).catch(error => {
        console.error("Error loading data file:", error);
    });
}

// Function to process template from template file
// This uses the data in the `meritData` JSON object and puts them into relevant cells in the template file
// A new excel file is generated and populated with data for every homeroom in every form
function processExcelTemplate(data) {

    const workbook = new ExcelJS.Workbook();
    workbook.xlsx.load(data).then(workbook => {
    
        console.log("Template file loaded successfully:", workbook);

        // TODO: Check if the uploaded template file is correct or not
        // Same as above, this can be done by checking if the file has a certain piece of text

        async function generateStatementsForm1() {
            for (let i = 0; i < 15; i++) {

                // Get the homeroom name
                let homeroom = meritData.form1.part1.data[i][1];

                const worksheet = workbook.getWorksheet('Sheet1');

                // Change the homeroom name in the template to the current one being iterated
                const cell = worksheet.getCell('B8');
                cell.value = `HOMEROOM 1${homeroom} ${year}`;

                // Part 1 

                let cellpt1_1 = worksheet.getCell('D16');
                cellpt1_1.value = meritData.form1.part1.data[i][3] || 0;

                let cellpt1_2 = worksheet.getCell('E16');
                cellpt1_2.value = meritData.form1.part1.data[i][4] || 0;

                let cellpt1_3 = worksheet.getCell('D17');
                cellpt1_3.value = meritData.form1.part1.data[i][5] || 0;

                let cellpt1_4 = worksheet.getCell('E17');
                cellpt1_4.value = meritData.form1.part1.data[i][6] || 0;

                let cellpt1_5 = worksheet.getCell('D18');
                cellpt1_5.value = meritData.form1.part1.data[i][7] || 0;

                let cellpt1_6 = worksheet.getCell('E18');
                cellpt1_6.value = meritData.form1.part1.data[i][8] || 0;

                let cellpt1_7 = worksheet.getCell('D19')
                cellpt1_7.value = meritData.form1.part1.data[i][9] || 0;

                let cellpt1_8 = worksheet.getCell('E19')
                cellpt1_8.value = meritData.form1.part1.data[i][10] || 0;
                

                // Part 2 (Cells D22, E22, D23, E23, D24, E24, D25, E25)

                let cellpt2_1 = worksheet.getCell('D22');
                cellpt2_1.value = meritData.form1.part2.data[i][0] || 0;

                let cellpt2_2 = worksheet.getCell('E22');
                cellpt2_2.value = meritData.form1.part2.data[i][1] || 0;

                let cellpt2_3 = worksheet.getCell('D23');
                cellpt2_3.value = meritData.form1.part2.data[i][2] || 0;

                let cellpt2_4 = worksheet.getCell('E23');
                cellpt2_4.value = meritData.form1.part2.data[i][3] || 0;

                let cellpt2_5 = worksheet.getCell('D24');
                cellpt2_5.value = meritData.form1.part2.data[i][4] || 0;

                let cellpt2_6 = worksheet.getCell('E24');
                cellpt2_6.value = meritData.form1.part2.data[i][5] || 0;

                let cellpt2_7 = worksheet.getCell('D25');
                cellpt2_7.value = meritData.form1.part2.data[i][6] || 0;

                let cellpt2_8 = worksheet.getCell('E25');
                cellpt2_8.value = meritData.form1.part2.data[i][7] || 0;

                // Part 3 (Cells E28, E29, E30, E31, )

                let cellpt3_1 = worksheet.getCell('E28');
                cellpt3_1.value = meritData.form1.part3.data[i][0] + meritData.form1.part3.data[i][1] + meritData.form1.part3.data[i][2] || 0;

                let cellpt3_2 = worksheet.getCell('E29');
                cellpt3_2.value = meritData.form1.part3.data[i][3] + meritData.form1.part3.data[i][4] + meritData.form1.part3.data[i][5] || 0;

                let cellpt3_3 = worksheet.getCell('E30');
                cellpt3_3.value = meritData.form1.part3.data[i][6] + meritData.form1.part3.data[i][7] + meritData.form1.part3.data[i][8] || 0;

                let cellpt3_4 = worksheet.getCell('E31');
                cellpt3_4.value = 0

                // Part 4 (Cells D34, E34, D35, E35, D36, E36, D37, E37, D38, E38, D39, E39, D40, E40) depending on how many pertandingans there are
                // Check number of pertandingans and only write to their respective cells, going down a row for each pertandingan

                let cellpt4_1 = worksheet.getCell('D34');
                cellpt4_1.value = meritData.form1.part4.data[i + 1][0] || 0;

                let cellpt4_2 = worksheet.getCell('E34');
                cellpt4_2.value = meritData.form1.part4.data[i + 1][1] || 0;

                let cellpt4_3 = worksheet.getCell('D35');
                cellpt4_3.value = meritData.form1.part4.data[i + 1][2] || 0;

                let cellpt4_4 = worksheet.getCell('E35');
                cellpt4_4.value = meritData.form1.part4.data[i + 1][3] || 0;

                let cellpt4_5 = worksheet.getCell('D36');
                cellpt4_5.value = meritData.form1.part4.data[i + 1][4] || 0;

                let cellpt4_6 = worksheet.getCell('E36');
                cellpt4_6.value = meritData.form1.part4.data[i + 1][5] || 0;

                let cellpt4_7 = worksheet.getCell('D37');
                cellpt4_7.value = meritData.form1.part4.data[i + 1][6] || 0;

                let cellpt4_8 = worksheet.getCell('E37');
                cellpt4_8.value = meritData.form1.part4.data[i + 1][7] || 0;

                let cellpt4_9 = worksheet.getCell('D38');
                cellpt4_9.value = meritData.form1.part4.data[i + 1][8] || 0;

                let cellpt4_10 = worksheet.getCell('E38');
                cellpt4_10.value = meritData.form1.part4.data[i + 1][9] || 0;

                let cellpt4_11 = worksheet.getCell('D39');
                cellpt4_11.value = meritData.form1.part4.data[i + 1][10] || 0;

                let cellpt4_12 = worksheet.getCell('E39');
                cellpt4_12.value = meritData.form1.part4.data[i + 1][11] || 0;

                let cellpt4_13 = worksheet.getCell('D40');
                cellpt4_13.value = meritData.form1.part4.data[i + 1][12] || 0;

                let cellpt4_14 = worksheet.getCell('E40');
                cellpt4_14.value = meritData.form1.part4.data[i + 1][13] || 0;

                let cellpt4_15 = worksheet.getCell('D41');
                cellpt4_15.value = meritData.form1.part4.data[i + 1][14] || 0;

                let cellpt4_16 = worksheet.getCell('E41');
                cellpt4_16.value = meritData.form1.part4.data[i + 1][15] || 0;

                let cellpt4_17 = worksheet.getCell('D42');
                cellpt4_17.value = meritData.form1.part4.data[i + 1][16] || 0;

                let cellpt4_18 = worksheet.getCell('E42');
                cellpt4_18.value = meritData.form1.part4.data[i + 1][17] || 0;

                let cellpt4_19 = worksheet.getCell('D43');
                cellpt4_19.value = meritData.form1.part4.data[i + 1][18] || 0;

                let cellpt4_20 = worksheet.getCell('E43');
                cellpt4_20.value = meritData.form1.part4.data[i + 1][19] || 0;

                let cellpt4_21 = worksheet.getCell('D44');
                cellpt4_21.value = meritData.form1.part4.data[i + 1][20] || 0;

                let cellpt4_22 = worksheet.getCell('E44');
                cellpt4_22.value = meritData.form1.part4.data[i + 1][21] || 0;

                let cellpt4_23 = worksheet.getCell('D45');
                cellpt4_23.value = meritData.form1.part4.data[i + 1][22] || 0;

                let cellpt4_24 = worksheet.getCell('E45');
                cellpt4_24.value = meritData.form1.part4.data[i + 1][23] || 0;

                let form1PertandinganNames = meritData.form1.part4.data[0];

                // Fill in the pertandingan names from C34 until C40

                for (let j = 0; j < form1PertandinganNames.length; j++) {
                    let cellpt4_9 = worksheet.getCell(`C${34 + j}`);
                    cellpt4_9.value = form1PertandinganNames[j];

                    let cellpt4_10 = worksheet.getCell('C41');
                    cellpt4_10.value = '';
                }

                // Push the completed excel file to the `statementFiles` array
                const buffer = await workbook.xlsx.writeBuffer();
                statementFiles.push({ name: `1${homeroom}-${year}.xlsx`, content: buffer });

            }
            
        }
        
        generateStatementsForm1()

        async function generateStatementsForm2() {
            for (let i = 0; i < 15; i++) {

                // Get the homeroom name
                let homeroom = meritData.form2.part1.data[i][1];

                const worksheet = workbook.getWorksheet('Sheet1');

                // Change the homeroom name in the template to the current homeroom name
                const cell = worksheet.getCell('B8');
                cell.value = `HOMEROOM 2${homeroom} ${year}`;

                // Part 1

                let cellpt1_1 = worksheet.getCell('D16');
                cellpt1_1.value = meritData.form2.part1.data[i][3] || 0;

                let cellpt1_2 = worksheet.getCell('E16');
                cellpt1_2.value = meritData.form2.part1.data[i][4] || 0;

                let cellpt1_3 = worksheet.getCell('D17');
                cellpt1_3.value = meritData.form2.part1.data[i][5] || 0;

                let cellpt1_4 = worksheet.getCell('E17');
                cellpt1_4.value = meritData.form2.part1.data[i][6] || 0;

                let cellpt1_5 = worksheet.getCell('D18');
                cellpt1_5.value = meritData.form2.part1.data[i][7] || 0;

                let cellpt1_6 = worksheet.getCell('E18');
                cellpt1_6.value = meritData.form2.part1.data[i][8] || 0;

                let cellpt1_7 = worksheet.getCell('D19')
                cellpt1_7.value = meritData.form2.part1.data[i][9] || 0;

                let cellpt1_8 = worksheet.getCell('E19')
                cellpt1_8.value = meritData.form2.part1.data[i][10] || 0;

                // Part 2 (Cells D22, E22, D23, E23, D24, E24, D25, E25)

                let cellpt2_1 = worksheet.getCell('D22');
                cellpt2_1.value = meritData.form2.part2.data[i][0] || 0;

                let cellpt2_2 = worksheet.getCell('E22');
                cellpt2_2.value = meritData.form2.part2.data[i][1] || 0;

                let cellpt2_3 = worksheet.getCell('D23');
                cellpt2_3.value = meritData.form2.part2.data[i][2] || 0;

                let cellpt2_4 = worksheet.getCell('E23');
                cellpt2_4.value = meritData.form2.part2.data[i][3] || 0;

                let cellpt2_5 = worksheet.getCell('D24');
                cellpt2_5.value = meritData.form2.part2.data[i][4] || 0;

                let cellpt2_6 = worksheet.getCell('E24');
                cellpt2_6.value = meritData.form2.part2.data[i][5] || 0;

                let cellpt2_7 = worksheet.getCell('D25');
                cellpt2_7.value = meritData.form2.part2.data[i][6] || 0;

                let cellpt2_8 = worksheet.getCell('E25');
                cellpt2_8.value = meritData.form2.part2.data[i][7] || 0;

                // Part 3 (Cells E28, E29, E30, E31, )

                let cellpt3_1 = worksheet.getCell('E28');
                cellpt3_1.value = meritData.form2.part3.data[i][0] + meritData.form2.part3.data[i][1] + meritData.form2.part3.data[i][2] || 0;

                let cellpt3_2 = worksheet.getCell('E29');
                cellpt3_2.value = meritData.form2.part3.data[i][3] + meritData.form2.part3.data[i][4] + meritData.form2.part3.data[i][5] || 0;

                let cellpt3_3 = worksheet.getCell('E30');
                cellpt3_3.value = meritData.form2.part3.data[i][6] + meritData.form2.part3.data[i][7] + meritData.form2.part3.data[i][8] || 0;

                let cellpt3_4 = worksheet.getCell('E31');
                cellpt3_4.value = 0

                // Part 4 (Cells D34, E34, D35, E35, D36, E36, D37, E37, D38, E38, D39, E39, D40, E40) depending on how many pertandingans there are
                // Check number of pertandingans and only write to their respective cells, going down for each pertandingan

                let cellpt4_1 = worksheet.getCell('D34');
                cellpt4_1.value = meritData.form2.part4.data[i + 1][0] || 0;

                let cellpt4_2 = worksheet.getCell('E34');
                cellpt4_2.value = meritData.form2.part4.data[i + 1][1] || 0;

                let cellpt4_3 = worksheet.getCell('D35');
                cellpt4_3.value = meritData.form2.part4.data[i + 1][2] || 0;

                let cellpt4_4 = worksheet.getCell('E35');
                cellpt4_4.value = meritData.form2.part4.data[i + 1][3] || 0;

                let cellpt4_5 = worksheet.getCell('D36');
                cellpt4_5.value = meritData.form2.part4.data[i + 1][4] || 0;

                let cellpt4_6 = worksheet.getCell('E36');
                cellpt4_6.value = meritData.form2.part4.data[i + 1][5] || 0;

                let cellpt4_7 = worksheet.getCell('D37');
                cellpt4_7.value = meritData.form2.part4.data[i + 1][6] || 0;

                let cellpt4_8 = worksheet.getCell('E37');
                cellpt4_8.value = meritData.form2.part4.data[i + 1][7] || 0;

                let cellpt4_9 = worksheet.getCell('D38');
                cellpt4_9.value = meritData.form2.part4.data[i + 1][8] || 0;

                let cellpt4_10 = worksheet.getCell('E38');
                cellpt4_10.value = meritData.form2.part4.data[i + 1][9] || 0;

                let cellpt4_11 = worksheet.getCell('D39');
                cellpt4_11.value = meritData.form2.part4.data[i + 1][10] || 0;

                let cellpt4_12 = worksheet.getCell('E39');
                cellpt4_12.value = meritData.form2.part4.data[i + 1][11] || 0;

                let cellpt4_13 = worksheet.getCell('D40');
                cellpt4_13.value = meritData.form2.part4.data[i + 1][12] || 0;

                let cellpt4_14 = worksheet.getCell('E40');
                cellpt4_14.value = meritData.form2.part4.data[i + 1][13] || 0;

                let cellpt4_15 = worksheet.getCell('D41');
                cellpt4_15.value = meritData.form2.part4.data[i + 1][14] || 0;

                let cellpt4_16 = worksheet.getCell('E41');
                cellpt4_16.value = meritData.form2.part4.data[i + 1][15] || 0;

                let cellpt4_17 = worksheet.getCell('D42');
                cellpt4_17.value = meritData.form2.part4.data[i + 1][16] || 0;

                let cellpt4_18 = worksheet.getCell('E42');
                cellpt4_18.value = meritData.form2.part4.data[i + 1][17] || 0;

                let cellpt4_19 = worksheet.getCell('D43');
                cellpt4_19.value = meritData.form2.part4.data[i + 1][18] || 0;

                let cellpt4_20 = worksheet.getCell('E43');
                cellpt4_20.value = meritData.form2.part4.data[i + 1][19] || 0;

                let cellpt4_21 = worksheet.getCell('D44');
                cellpt4_21.value = meritData.form2.part4.data[i + 1][20] || 0;

                let cellpt4_22 = worksheet.getCell('E44');
                cellpt4_22.value = meritData.form2.part4.data[i + 1][21] || 0;

                let cellpt4_23 = worksheet.getCell('D45');
                cellpt4_23.value = meritData.form2.part4.data[i + 1][22] || 0;

                let cellpt4_24 = worksheet.getCell('E45');
                cellpt4_24.value = meritData.form2.part4.data[i + 1][23] || 0;

                let form2PertandinganNames = meritData.form2.part4.data[0];

                // Fill in the pertandingan names from C34 until C40

                for (let j = 0; j < form2PertandinganNames.length; j++) {
                    let cellpt4_9 = worksheet.getCell(`C${34 + j}`);
                    cellpt4_9.value = form2PertandinganNames[j];

                    let cellpt4_10 = worksheet.getCell('C41');
                    cellpt4_10.value = '';
                }

                // Push the completed excel file to the `statementFiles` array
                const buffer = await workbook.xlsx.writeBuffer();
                statementFiles.push({ name: `2${homeroom}-${year}.xlsx`, content: buffer });
                
            }
            
        }
        
        generateStatementsForm2()

        async function generateStatementsForm3() {
            for (let i = 0; i < 15; i++) {

                // Get the homeroom name
                let homeroom = meritData.form3.part1.data[i][1];

                const worksheet = workbook.getWorksheet('Sheet1');

                // Change the homeroom name in the template to the currrent one
                const cell = worksheet.getCell('B8');
                cell.value = `HOMEROOM 3${homeroom} ${year}`;

                // Part 1

                let cellpt1_1 = worksheet.getCell('D16');
                cellpt1_1.value = meritData.form3.part1.data[i][3] || 0;

                let cellpt1_2 = worksheet.getCell('E16');
                cellpt1_2.value = meritData.form3.part1.data[i][4] || 0;

                let cellpt1_3 = worksheet.getCell('D17');
                cellpt1_3.value = meritData.form3.part1.data[i][5] || 0;

                let cellpt1_4 = worksheet.getCell('E17');
                cellpt1_4.value = meritData.form3.part1.data[i][6] || 0;

                let cellpt1_5 = worksheet.getCell('D18');
                cellpt1_5.value = meritData.form3.part1.data[i][7] || 0;

                let cellpt1_6 = worksheet.getCell('E18');
                cellpt1_6.value = meritData.form3.part1.data[i][8] || 0;

                let cellpt1_7 = worksheet.getCell('D19')
                cellpt1_7.value = meritData.form3.part1.data[i][9] || 0;

                let cellpt1_8 = worksheet.getCell('E19')
                cellpt1_8.value = meritData.form3.part1.data[i][10] || 0;

                // Part 2 (Cells D22, E22, D23, E23, D24, E24, D25, E25)

                let cellpt2_1 = worksheet.getCell('D22');
                cellpt2_1.value = meritData.form3.part2.data[i][0] || 0;

                let cellpt2_2 = worksheet.getCell('E22');
                cellpt2_2.value = meritData.form3.part2.data[i][1] || 0;

                let cellpt2_3 = worksheet.getCell('D23');
                cellpt2_3.value = meritData.form3.part2.data[i][2] || 0;

                let cellpt2_4 = worksheet.getCell('E23');
                cellpt2_4.value = meritData.form3.part2.data[i][3] || 0;

                let cellpt2_5 = worksheet.getCell('D24');
                cellpt2_5.value = meritData.form3.part2.data[i][4] || 0;

                let cellpt2_6 = worksheet.getCell('E24');
                cellpt2_6.value = meritData.form3.part2.data[i][5] || 0;

                let cellpt2_7 = worksheet.getCell('D25');
                cellpt2_7.value = meritData.form3.part2.data[i][6] || 0;

                let cellpt2_8 = worksheet.getCell('E25');
                cellpt2_8.value = meritData.form3.part2.data[i][7] || 0;

                // Part 3 (Cells E28, E29, E30, E31, )

                let cellpt3_1 = worksheet.getCell('E28');
                cellpt3_1.value = meritData.form3.part3.data[i][0] + meritData.form3.part3.data[i][1] + meritData.form3.part3.data[i][2] || 0;

                let cellpt3_2 = worksheet.getCell('E29');
                cellpt3_2.value = meritData.form3.part3.data[i][3] + meritData.form3.part3.data[i][4] + meritData.form3.part3.data[i][5] || 0;

                let cellpt3_3 = worksheet.getCell('E30');
                cellpt3_3.value = meritData.form3.part3.data[i][6] + meritData.form3.part3.data[i][7] + meritData.form3.part3.data[i][8] || 0;

                let cellpt3_4 = worksheet.getCell('E31');
                cellpt3_4.value = 0

                // Part 4 (Cells D34, E34, D35, E35, D36, E36, D37, E37, D38, E38, D39, E39, D40, E40) depending on how many pertandingans there are
                // Check number of pertandingans and only write to their respective cells, going down for each pertandingan

                let cellpt4_1 = worksheet.getCell('D34');
                cellpt4_1.value = meritData.form3.part4.data[i + 1][0] || 0;

                let cellpt4_2 = worksheet.getCell('E34');
                cellpt4_2.value = meritData.form3.part4.data[i + 1][1] || 0;

                let cellpt4_3 = worksheet.getCell('D35');
                cellpt4_3.value = meritData.form3.part4.data[i + 1][2] || 0;

                let cellpt4_4 = worksheet.getCell('E35');
                cellpt4_4.value = meritData.form3.part4.data[i + 1][3] || 0;

                let cellpt4_5 = worksheet.getCell('D36');
                cellpt4_5.value = meritData.form3.part4.data[i + 1][4] || 0;

                let cellpt4_6 = worksheet.getCell('E36');
                cellpt4_6.value = meritData.form3.part4.data[i + 1][5] || 0;

                let cellpt4_7 = worksheet.getCell('D37');
                cellpt4_7.value = meritData.form3.part4.data[i + 1][6] || 0;

                let cellpt4_8 = worksheet.getCell('E37');
                cellpt4_8.value = meritData.form3.part4.data[i + 1][7] || 0;

                let cellpt4_9 = worksheet.getCell('D38');
                cellpt4_9.value = meritData.form3.part4.data[i + 1][8] || 0;

                let cellpt4_10 = worksheet.getCell('E38');
                cellpt4_10.value = meritData.form3.part4.data[i + 1][9] || 0;

                let cellpt4_11 = worksheet.getCell('D39');
                cellpt4_11.value = meritData.form3.part4.data[i + 1][10] || 0;

                let cellpt4_12 = worksheet.getCell('E39');
                cellpt4_12.value = meritData.form3.part4.data[i + 1][11] || 0;

                let cellpt4_13 = worksheet.getCell('D40');
                cellpt4_13.value = meritData.form3.part4.data[i + 1][12] || 0;

                let cellpt4_14 = worksheet.getCell('E40');
                cellpt4_14.value = meritData.form3.part4.data[i + 1][13] || 0;

                let cellpt4_15 = worksheet.getCell('D41');
                cellpt4_15.value = meritData.form3.part4.data[i + 1][14] || 0;

                let cellpt4_16 = worksheet.getCell('E41');
                cellpt4_16.value = meritData.form3.part4.data[i + 1][15] || 0;

                let cellpt4_17 = worksheet.getCell('D42');
                cellpt4_17.value = meritData.form3.part4.data[i + 1][16] || 0;

                let cellpt4_18 = worksheet.getCell('E42');
                cellpt4_18.value = meritData.form3.part4.data[i + 1][17] || 0;

                let cellpt4_19 = worksheet.getCell('D43');
                cellpt4_19.value = meritData.form3.part4.data[i + 1][18] || 0;

                let cellpt4_20 = worksheet.getCell('E43');
                cellpt4_20.value = meritData.form3.part4.data[i + 1][19] || 0;

                let cellpt4_21 = worksheet.getCell('D44');
                cellpt4_21.value = meritData.form3.part4.data[i + 1][20] || 0;

                let cellpt4_22 = worksheet.getCell('E44');
                cellpt4_22.value = meritData.form3.part4.data[i + 1][21] || 0;

                let cellpt4_23 = worksheet.getCell('D45');
                cellpt4_23.value = meritData.form3.part4.data[i + 1][22] || 0;

                let cellpt4_24 = worksheet.getCell('E45');
                cellpt4_24.value = meritData.form3.part4.data[i + 1][23] || 0;

                let form3PertandinganNames = meritData.form3.part4.data[0];

                // Fill in the pertandingan names from C34 until C40

                for (let j = 0; j < form3PertandinganNames.length; j++) {
                    let cellpt4_9 = worksheet.getCell(`C${34 + j}`);
                    cellpt4_9.value = form3PertandinganNames[j];

                    let cellpt4_10 = worksheet.getCell('C41');
                    cellpt4_10.value = '';
                }

                // Push the generated excel files to the `statementFiles` array
                const buffer = await workbook.xlsx.writeBuffer();
                statementFiles.push({ name: `3${homeroom}-${year}.xlsx`, content: buffer });
                
            }
            
        }
        
        generateStatementsForm3()

        async function generateStatementsForm4() {
            for (let i = 0; i < 15; i++) {
                
                // Get the homeroom name
                let homeroom = meritData.form4.part1.data[i][1];

                const worksheet = workbook.getWorksheet('Sheet1');

                // Change the homeroom name in the template to the current homeroom
                const cell = worksheet.getCell('B8');
                cell.value = `HOMEROOM 4${homeroom} ${year}`;

                // Part 1

                let cellpt1_1 = worksheet.getCell('D16');
                cellpt1_1.value = meritData.form4.part1.data[i][3] || 0;

                let cellpt1_2 = worksheet.getCell('E16');
                cellpt1_2.value = meritData.form4.part1.data[i][4] || 0;

                let cellpt1_3 = worksheet.getCell('D17');
                cellpt1_3.value = meritData.form4.part1.data[i][5] || 0;

                let cellpt1_4 = worksheet.getCell('E17');
                cellpt1_4.value = meritData.form4.part1.data[i][6] || 0;

                let cellpt1_5 = worksheet.getCell('D18');
                cellpt1_5.value = meritData.form4.part1.data[i][7] || 0;

                let cellpt1_6 = worksheet.getCell('E18');
                cellpt1_6.value = meritData.form4.part1.data[i][8] || 0;

                let cellpt1_7 = worksheet.getCell('D19')
                cellpt1_7.value = meritData.form4.part1.data[i][9] || 0;

                let cellpt1_8 = worksheet.getCell('E19')
                cellpt1_8.value = meritData.form4.part1.data[i][10] || 0;

                // Part 2 (Cells D22, E22, D23, E23, D24, E24, D25, E25)

                let cellpt2_1 = worksheet.getCell('D22');
                cellpt2_1.value = meritData.form4.part2.data[i][0] || 0;

                let cellpt2_2 = worksheet.getCell('E22');
                cellpt2_2.value = meritData.form4.part2.data[i][1] || 0;

                let cellpt2_3 = worksheet.getCell('D23');
                cellpt2_3.value = meritData.form4.part2.data[i][2] || 0;

                let cellpt2_4 = worksheet.getCell('E23');
                cellpt2_4.value = meritData.form4.part2.data[i][3] || 0;

                let cellpt2_5 = worksheet.getCell('D24');
                cellpt2_5.value = meritData.form4.part2.data[i][4] || 0;

                let cellpt2_6 = worksheet.getCell('E24');
                cellpt2_6.value = meritData.form4.part2.data[i][5] || 0;

                let cellpt2_7 = worksheet.getCell('D25');
                cellpt2_7.value = meritData.form4.part2.data[i][6] || 0;

                let cellpt2_8 = worksheet.getCell('E25');
                cellpt2_8.value = meritData.form4.part2.data[i][7] || 0;

                // Part 3 (Cells E28, E29, E30, E31, )

                let cellpt3_1 = worksheet.getCell('E28');
                cellpt3_1.value = meritData.form4.part3.data[i][0] + meritData.form4.part3.data[i][1] + meritData.form4.part3.data[i][2] || 0;

                let cellpt3_2 = worksheet.getCell('E29');
                cellpt3_2.value = meritData.form4.part3.data[i][3] + meritData.form4.part3.data[i][4] + meritData.form4.part3.data[i][5] || 0;

                let cellpt3_3 = worksheet.getCell('E30');
                cellpt3_3.value = meritData.form4.part3.data[i][6] + meritData.form4.part3.data[i][7] + meritData.form4.part3.data[i][8] || 0;

                let cellpt3_4 = worksheet.getCell('E31');
                cellpt3_4.value = 0

                // Part 4 (Cells D34, E34, D35, E35, D36, E36, D37, E37, D38, E38, D39, E39, D40, E40) depending on how many pertandingans there are
                // Check number of pertandingans and only write to their respective cells, going down for each pertandingan

                let cellpt4_1 = worksheet.getCell('D34');
                cellpt4_1.value = meritData.form4.part4.data[i + 1][0] || 0;

                let cellpt4_2 = worksheet.getCell('E34');
                cellpt4_2.value = meritData.form4.part4.data[i + 1][1] || 0;

                let cellpt4_3 = worksheet.getCell('D35');
                cellpt4_3.value = meritData.form4.part4.data[i + 1][2] || 0;

                let cellpt4_4 = worksheet.getCell('E35');
                cellpt4_4.value = meritData.form4.part4.data[i + 1][3] || 0;

                let cellpt4_5 = worksheet.getCell('D36');
                cellpt4_5.value = meritData.form4.part4.data[i + 1][4] || 0;

                let cellpt4_6 = worksheet.getCell('E36');
                cellpt4_6.value = meritData.form4.part4.data[i + 1][5] || 0;

                let cellpt4_7 = worksheet.getCell('D37');
                cellpt4_7.value = meritData.form4.part4.data[i + 1][6] || 0;

                let cellpt4_8 = worksheet.getCell('E37');
                cellpt4_8.value = meritData.form4.part4.data[i + 1][7] || 0;

                let cellpt4_9 = worksheet.getCell('D38');
                cellpt4_9.value = meritData.form4.part4.data[i + 1][8] || 0;

                let cellpt4_10 = worksheet.getCell('E38');
                cellpt4_10.value = meritData.form4.part4.data[i + 1][9] || 0;

                let cellpt4_11 = worksheet.getCell('D39');
                cellpt4_11.value = meritData.form4.part4.data[i + 1][10] || 0;

                let cellpt4_12 = worksheet.getCell('E39');
                cellpt4_12.value = meritData.form4.part4.data[i + 1][11] || 0;

                let cellpt4_13 = worksheet.getCell('D40');
                cellpt4_13.value = meritData.form4.part4.data[i + 1][12] || 0;

                let cellpt4_14 = worksheet.getCell('E40');
                cellpt4_14.value = meritData.form4.part4.data[i + 1][13] || 0;

                let cellpt4_15 = worksheet.getCell('D41');
                cellpt4_15.value = meritData.form4.part4.data[i + 1][14] || 0;

                let cellpt4_16 = worksheet.getCell('E41');
                cellpt4_16.value = meritData.form4.part4.data[i + 1][15] || 0;

                let cellpt4_17 = worksheet.getCell('D42');
                cellpt4_17.value = meritData.form4.part4.data[i + 1][16] || 0;

                let cellpt4_18 = worksheet.getCell('E42');
                cellpt4_18.value = meritData.form4.part4.data[i + 1][17] || 0;

                let cellpt4_19 = worksheet.getCell('D43');
                cellpt4_19.value = meritData.form4.part4.data[i + 1][18] || 0;

                let cellpt4_20 = worksheet.getCell('E43');
                cellpt4_20.value = meritData.form4.part4.data[i + 1][19] || 0;

                let cellpt4_21 = worksheet.getCell('D44');
                cellpt4_21.value = meritData.form4.part4.data[i + 1][20] || 0;

                let cellpt4_22 = worksheet.getCell('E44');
                cellpt4_22.value = meritData.form4.part4.data[i + 1][21] || 0;

                let cellpt4_23 = worksheet.getCell('D45');
                cellpt4_23.value = meritData.form4.part4.data[i + 1][22] || 0;

                let cellpt4_24 = worksheet.getCell('E45');
                cellpt4_24.value = meritData.form4.part4.data[i + 1][23] || 0;

                let form4PertandinganNames = meritData.form4.part4.data[0];

                // Fill in the pertandingan names from C34 until C40
                for (let j = 0; j < form4PertandinganNames.length; j++) {
                    let cellpt4_9 = worksheet.getCell(`C${34 + j}`);
                    cellpt4_9.value = form4PertandinganNames[j];

                    let cellpt4_10 = worksheet.getCell('C41');
                    cellpt4_10.value = '';
                }

                // Push the generated excel file to the `statementFiles` array
                const buffer = await workbook.xlsx.writeBuffer();
                statementFiles.push({ name: `4${homeroom}-${year}.xlsx`, content: buffer });
                
            }
            
        }
        
        generateStatementsForm4()

        async function generateStatementsForm5() {
            for (let i = 0; i < 15; i++) {
                
                // Get the homeroom name
                let homeroom = meritData.form5.part1.data[i][1];

                const worksheet = workbook.getWorksheet('Sheet1');

                // Change the homeroom name
                const cell = worksheet.getCell('B8');
                cell.value = `HOMEROOM 5${homeroom} ${year}`;

                // Part 1

                let cellpt1_1 = worksheet.getCell('D16');
                cellpt1_1.value = meritData.form5.part1.data[i][3] || 0;

                let cellpt1_2 = worksheet.getCell('E16');
                cellpt1_2.value = meritData.form5.part1.data[i][4] || 0;

                let cellpt1_3 = worksheet.getCell('D17');
                cellpt1_3.value = meritData.form5.part1.data[i][5] || 0;

                let cellpt1_4 = worksheet.getCell('E17');
                cellpt1_4.value = meritData.form5.part1.data[i][6] || 0;

                let cellpt1_5 = worksheet.getCell('D18');
                cellpt1_5.value = meritData.form5.part1.data[i][7] || 0;

                let cellpt1_6 = worksheet.getCell('E18');
                cellpt1_6.value = meritData.form5.part1.data[i][8] || 0;

                let cellpt1_7 = worksheet.getCell('D19')
                cellpt1_7.value = meritData.form5.part1.data[i][9] || 0;

                let cellpt1_8 = worksheet.getCell('E19')
                cellpt1_8.value = meritData.form5.part1.data[i][10] || 0;

                // Part 2 (Cells D22, E22, D23, E23, D24, E24, D25, E25)

                let cellpt2_1 = worksheet.getCell('D22');
                cellpt2_1.value = meritData.form5.part2.data[i][0] || 0;

                let cellpt2_2 = worksheet.getCell('E22');
                cellpt2_2.value = meritData.form5.part2.data[i][1] || 0;

                let cellpt2_3 = worksheet.getCell('D23');
                cellpt2_3.value = meritData.form5.part2.data[i][2] || 0;

                let cellpt2_4 = worksheet.getCell('E23');
                cellpt2_4.value = meritData.form5.part2.data[i][3] || 0;

                let cellpt2_5 = worksheet.getCell('D24');
                cellpt2_5.value = meritData.form5.part2.data[i][4] || 0;

                let cellpt2_6 = worksheet.getCell('E24');
                cellpt2_6.value = meritData.form5.part2.data[i][5] || 0;

                let cellpt2_7 = worksheet.getCell('D25');
                cellpt2_7.value = meritData.form5.part2.data[i][6] || 0;

                let cellpt2_8 = worksheet.getCell('E25');
                cellpt2_8.value = meritData.form5.part2.data[i][7] || 0;

                // Part 3 (Cells E28, E29, E30, E31, )

                let cellpt3_1 = worksheet.getCell('E28');
                cellpt3_1.value = meritData.form5.part3.data[i][0] + meritData.form5.part3.data[i][1] + meritData.form5.part3.data[i][2] || 0;

                let cellpt3_2 = worksheet.getCell('E29');
                cellpt3_2.value = meritData.form5.part3.data[i][3] + meritData.form5.part3.data[i][4] + meritData.form5.part3.data[i][5] || 0;

                let cellpt3_3 = worksheet.getCell('E30');
                cellpt3_3.value = meritData.form5.part3.data[i][6] + meritData.form5.part3.data[i][7] + meritData.form5.part3.data[i][8] || 0;

                let cellpt3_4 = worksheet.getCell('E31');
                cellpt3_4.value = 0

                // Part 4 (Cells D34, E34, D35, E35, D36, E36, D37, E37, D38, E38, D39, E39, D40, E40) depending on how many pertandingans there are
                // Check number of pertandingans and only write to their respective cells, going down for each pertandingan

                let cellpt4_1 = worksheet.getCell('D34');
                cellpt4_1.value = meritData.form5.part4.data[i + 1][0] || 0;

                let cellpt4_2 = worksheet.getCell('E34');
                cellpt4_2.value = meritData.form5.part4.data[i + 1][1] || 0;

                let cellpt4_3 = worksheet.getCell('D35');
                cellpt4_3.value = meritData.form5.part4.data[i + 1][2] || 0;

                let cellpt4_4 = worksheet.getCell('E35');
                cellpt4_4.value = meritData.form5.part4.data[i + 1][3] || 0;

                let cellpt4_5 = worksheet.getCell('D36');
                cellpt4_5.value = meritData.form5.part4.data[i + 1][4] || 0;

                let cellpt4_6 = worksheet.getCell('E36');
                cellpt4_6.value = meritData.form5.part4.data[i + 1][5] || 0;

                let cellpt4_7 = worksheet.getCell('D37');
                cellpt4_7.value = meritData.form5.part4.data[i + 1][6] || 0;

                let cellpt4_8 = worksheet.getCell('E37');
                cellpt4_8.value = meritData.form5.part4.data[i + 1][7] || 0;

                let cellpt4_9 = worksheet.getCell('D38');
                cellpt4_9.value = meritData.form5.part4.data[i + 1][8] || 0;

                let cellpt4_10 = worksheet.getCell('E38');
                cellpt4_10.value = meritData.form5.part4.data[i + 1][9] || 0;

                let cellpt4_11 = worksheet.getCell('D39');
                cellpt4_11.value = meritData.form5.part4.data[i + 1][10] || 0;

                let cellpt4_12 = worksheet.getCell('E39');
                cellpt4_12.value = meritData.form5.part4.data[i + 1][11] || 0;

                let cellpt4_13 = worksheet.getCell('D40');
                cellpt4_13.value = meritData.form5.part4.data[i + 1][12] || 0;

                let cellpt4_14 = worksheet.getCell('E40');
                cellpt4_14.value = meritData.form5.part4.data[i + 1][13] || 0;

                let cellpt4_15 = worksheet.getCell('D41');
                cellpt4_15.value = meritData.form5.part4.data[i + 1][14] || 0;

                let cellpt4_16 = worksheet.getCell('E41');
                cellpt4_16.value = meritData.form5.part4.data[i + 1][15] || 0;

                let cellpt4_17 = worksheet.getCell('D42');
                cellpt4_17.value = meritData.form5.part4.data[i + 1][16] || 0;

                let cellpt4_18 = worksheet.getCell('E42');
                cellpt4_18.value = meritData.form5.part4.data[i + 1][17] || 0;

                let cellpt4_19 = worksheet.getCell('D43');
                cellpt4_19.value = meritData.form5.part4.data[i + 1][18] || 0;

                let cellpt4_20 = worksheet.getCell('E43');
                cellpt4_20.value = meritData.form5.part4.data[i + 1][19] || 0;

                let cellpt4_21 = worksheet.getCell('D44');
                cellpt4_21.value = meritData.form5.part4.data[i + 1][20] || 0;

                let cellpt4_22 = worksheet.getCell('E44');
                cellpt4_22.value = meritData.form5.part4.data[i + 1][21] || 0;

                let cellpt4_23 = worksheet.getCell('D45');
                cellpt4_23.value = meritData.form5.part4.data[i + 1][22] || 0;

                let cellpt4_24 = worksheet.getCell('E45');
                cellpt4_24.value = meritData.form5.part4.data[i + 1][23] || 0;

                let form5PertandinganNames = meritData.form5.part4.data[0];

                // Fill in the pertandingan names from C34 until C45
                for (let j = 0; j < form5PertandinganNames.length; j++) {
                    let cellpt4_9 = worksheet.getCell(`C${34 + j}`);
                    cellpt4_9.value = form5PertandinganNames[j];

                    let cellpt4_10 = worksheet.getCell('C45');
                    cellpt4_10.value = '';
                }

                // Write the data to the copy
                const buffer = await workbook.xlsx.writeBuffer();
                statementFiles.push({ name: `5${homeroom}-${year}.xlsx`, content: buffer });
                
            }
            
        }
        
        generateStatementsForm5()

    });

}

const downloadButton = document.getElementById('download')
downloadButton.addEventListener('click', generateAndDownloadZip);

async function createZip(files) {
    const zip = new JSZip();

    files.forEach(file => {
        zip.file(file.name, file.content);
    });

    // TODO: In the final zip file, segregate every form into their own folders

    const content = await zip.generateAsync({ type: 'blob' });
    saveAs(content, `penyata-merit-demerit-HR-${year}.zip`);
}

async function generateAndDownloadZip() {
    await createZip(statementFiles);
}