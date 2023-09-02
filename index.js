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

filename = 'merit-demerit-homeroom-2022.xlsx'

async function readExcelFileFromSystem(filename) {

    await workbook.xlsx.readFile(filename);

    readDataFromFile()

}

readExcelFileFromSystem(filename)

async function readDataFromFile() {
    form1Worksheet = workbook.getWorksheet('TING.1')
    form2Worksheet = workbook.getWorksheet('TING.2')
    form3Worksheet = workbook.getWorksheet('TING.3')
    form4Worksheet = workbook.getWorksheet('TING.4')
    form5Worksheet = workbook.getWorksheet('TING.5')

    // Logging all of the actualRowCounts of every form

    console.log("Actual row counts of every form:")
    console.log("Form 1: " + form1Worksheet.actualRowCount)
    console.log("Form 2: " + form2Worksheet.actualRowCount)
    console.log("Form 3: " + form3Worksheet.actualRowCount)
    console.log("Form 4: " + form4Worksheet.actualRowCount)
    console.log("Form 5: " + form5Worksheet.actualRowCount)

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
        if (item === 'NAMA PERTANDINGAN' || item === 'JUMLAH') {
            return null
        } else {
            return item
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

async function exportToJSON() {
    // Export all of the read data to a JSON file

    // Create a JSON object

    let data = {
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

    // Save the JSON object to a file

    fs.writeFile("data.json", JSON.stringify(data), (err) => {
        if (err) {
            console.log(err)
        } else {
            console.log("Successfully wrote to file")
        }
    }
    )
}
