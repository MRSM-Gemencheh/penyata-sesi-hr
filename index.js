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

    // If a form has 91 rows then the number of homerooms is 15
    // If a form has 106 rows then the number of homerooms is 18
    // This will affect what rows we'll need to get data from below

    // Part 1

    // For each form, if a form has 15 homerooms, then we'll need to get data from rows 6 to 20
    // For each form, if a form has 18 homerooms, then we'll need to get data from rows 6 to 23 
    // For each row, get the data for columns 3 to 11 only
    // Store the data in an array of arrays

    // Form 1

    form1Part1Data = []

    if (form1Worksheet.actualRowCount == 91) {
        for (let i = 6; i <= 20; i++) {
            form1Part1Data.push(form1Worksheet.getRow(i).values.slice(3, 12))
        }
    } else if (form1Worksheet.actualRowCount == 106) {
        for (let i = 6; i <= 23; i++) {
            form1Part1Data.push(form1Worksheet.getRow(i).values.slice(3, 12))
        }
    }

    // Form 2

    form2Part1Data = []

    if (form2Worksheet.actualRowCount == 91) {
        for (let i = 6; i <= 20; i++) {
            form2Part1Data.push(form2Worksheet.getRow(i).values.slice(3, 12))
        }
    } else if (form2Worksheet.actualRowCount == 106) {
        for (let i = 6; i <= 23; i++) {
            form2Part1Data.push(form2Worksheet.getRow(i).values.slice(3, 12))
        }
    }

    // Form 3

    form3Part1Data = []

    if (form3Worksheet.actualRowCount == 91) {
        for (let i = 6; i <= 20; i++) {
            form3Part1Data.push(form3Worksheet.getRow(i).values.slice(3, 12))
        }
    } else if (form3Worksheet.actualRowCount == 106) {
        for (let i = 6; i <= 23; i++) {
            form3Part1Data.push(form3Worksheet.getRow(i).values.slice(3, 12))
        }
    }

    // Form 4

    form4Part1Data = []

    if (form4Worksheet.actualRowCount == 91) {
        for (let i = 6; i <= 20; i++) {
            form4Part1Data.push(form4Worksheet.getRow(i).values.slice(3, 12))
        }
    } else if (form4Worksheet.actualRowCount == 106) {
        for (let i = 6; i <= 23; i++) {
            form4Part1Data.push(form4Worksheet.getRow(i).values.slice(3, 12))
        }
    }

    // Form 5

    form5Part1Data = []

    if (form5Worksheet.actualRowCount == 91) {
        for (let i = 6; i <= 20; i++) {
            form5Part1Data.push(form5Worksheet.getRow(i).values.slice(3, 12))
        }
    } else if (form5Worksheet.actualRowCount == 106) {
        for (let i = 6; i <= 23; i++) {
            form5Part1Data.push(form5Worksheet.getRow(i).values.slice(3, 12))
        }
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

    // For each form, if a form has 91 rows, then we'll need to get data from rows 26 to 40
    // For each form, if a form has 106 rows, then we'll need to get data from rows 29 to 46
    // For each row, get the data for columns 3 to 11 only
    // Store the data in an array of arrays

    // Form 1

    form1Part2Data = []

    if (form1Worksheet.actualRowCount == 91) {
        for (let i = 26; i <= 40; i++) {
            form1Part2Data.push(form1Worksheet.getRow(i).values.slice(3, 12))
        }
    } else if (form1Worksheet.actualRowCount == 106) {
        for (let i = 29; i <= 46; i++) {
            form1Part2Data.push(form1Worksheet.getRow(i).values.slice(3, 12))
        }
    }

    // Form 2

    form2Part2Data = []

    if (form2Worksheet.actualRowCount == 91) {
        for (let i = 26; i <= 40; i++) {
            form2Part2Data.push(form2Worksheet.getRow(i).values.slice(3, 12))
        }
    } else if (form2Worksheet.actualRowCount == 106) {
        for (let i = 29; i <= 46; i++) {
            form2Part2Data.push(form2Worksheet.getRow(i).values.slice(3, 12))
        }
    }

    // Form 3

    form3Part2Data = []

    if (form3Worksheet.actualRowCount == 91) {
        for (let i = 26; i <= 40; i++) {
            form3Part2Data.push(form3Worksheet.getRow(i).values.slice(3, 12))
        }
    } else if (form3Worksheet.actualRowCount == 106) {
        for (let i = 29; i <= 46; i++) {
            form3Part2Data.push(form3Worksheet.getRow(i).values.slice(3, 12))
        }
    }

    // Form 4

    form4Part2Data = []

    if (form4Worksheet.actualRowCount == 91) {
        for (let i = 26; i <= 40; i++) {
            form4Part2Data.push(form4Worksheet.getRow(i).values.slice(3, 12))
        }
    } else if (form4Worksheet.actualRowCount == 106) {
        for (let i = 29; i <= 46; i++) {
            form4Part2Data.push(form4Worksheet.getRow(i).values.slice(3, 12))
        }
    }

    // Form 5

    form5Part2Data = []

    if (form5Worksheet.actualRowCount == 91) {
        for (let i = 26; i <= 40; i++) {
            form5Part2Data.push(form5Worksheet.getRow(i).values.slice(3, 12))
        }
    } else if (form5Worksheet.actualRowCount == 106) {
        for (let i = 29; i <= 46; i++) {
            form5Part2Data.push(form5Worksheet.getRow(i).values.slice(3, 12))
        }
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

    // For each form, if a form has 91 rows, then we'll need to get data from rows 46 to 60
    // For each form, if a form has 106 rows, then we'll need to get data from rows 52 to 69
    // For each row, get the data for columns 3 to 12 only
    // Store the data in an array of arrays

    // Form 1

    form1Part3Data = []

    if (form1Worksheet.actualRowCount == 91) {
        for (let i = 46; i <= 60; i++) {
            form1Part3Data.push(form1Worksheet.getRow(i).values.slice(3, 13))
        }
    } else if (form1Worksheet.actualRowCount == 106) {
        for (let i = 52; i <= 69; i++) {
            form1Part3Data.push(form1Worksheet.getRow(i).values.slice(3, 13))
        }
    }

    // Form 2

    form2Part3Data = []

    if (form2Worksheet.actualRowCount == 91) {
        for (let i = 46; i <= 60; i++) {
            form2Part3Data.push(form2Worksheet.getRow(i).values.slice(3, 13))
        }
    } else if (form2Worksheet.actualRowCount == 106) {
        for (let i = 52; i <= 69; i++) {
            form2Part3Data.push(form2Worksheet.getRow(i).values.slice(3, 13))
        }
    }

    // Form 3

    form3Part3Data = []

    if (form3Worksheet.actualRowCount == 91) {
        for (let i = 46; i <= 60; i++) {
            form3Part3Data.push(form3Worksheet.getRow(i).values.slice(3, 13))
        }
    } else if (form3Worksheet.actualRowCount == 106) {
        for (let i = 52; i <= 69; i++) {
            form3Part3Data.push(form3Worksheet.getRow(i).values.slice(3, 13))
        }
    }

    // Form 4

    form4Part3Data = []

    if (form4Worksheet.actualRowCount == 91) {
        for (let i = 46; i <= 60; i++) {
            form4Part3Data.push(form4Worksheet.getRow(i).values.slice(3, 13))
        }
    } else if (form4Worksheet.actualRowCount == 106) {
        for (let i = 52; i <= 69; i++) {
            form4Part3Data.push(form4Worksheet.getRow(i).values.slice(3, 13))
        }
    }

    // Form 5

    form5Part3Data = []

    if (form5Worksheet.actualRowCount == 91) {
        for (let i = 46; i <= 60; i++) {
            form5Part3Data.push(form5Worksheet.getRow(i).values.slice(3, 13))
        }
    } else if (form5Worksheet.actualRowCount == 106) {
        for (let i = 52; i <= 69; i++) {
            form5Part3Data.push(form5Worksheet.getRow(i).values.slice(3, 13))
        }
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

    // For each form, if a form has 91 rows, then we'll need to get data from rows 67 to 81
    // For each form, if a form has 106 rows, then we'll need to get data from rows 77 to 94

    jumlahPertandinganForm1 = 4
    jumlahPertandinganForm2 = 5
    jumlahPertandinganForm3 = 5
    jumlahPertandinganForm4 = 4
    jumlahPertandinganForm5 = 5

    // Determining the number of columns to get data from based on the number of pertandingans
    // Multiply the number of each pertandingan by 2 then add 4 to determine the end column number
    // For example, if there are 4 pertandingans, then we'll need to get data from columns 3 to 11

    // Store the data in an array of arrays

    // Form 1

    form1Part4Data = []

    if (form1Worksheet.actualRowCount == 91) {
        for (let i = 67; i <= 81; i++) {
            form1Part4Data.push(form1Worksheet.getRow(i).values.slice(3, jumlahPertandinganForm1 * 2 + 4))
        }
    } else if (form1Worksheet.actualRowCount == 106) {
        for (let i = 77; i <= 94; i++) {
            form1Part4Data.push(form1Worksheet.getRow(i).values.slice(3, jumlahPertandinganForm1 * 2 + 4))
        }
    }

    // Form 2

    form2Part4Data = []

    if (form2Worksheet.actualRowCount == 91) {
        for (let i = 67; i <= 81; i++) {
            form2Part4Data.push(form2Worksheet.getRow(i).values.slice(3, jumlahPertandinganForm2 * 2 + 4))
        }
    } else if (form2Worksheet.actualRowCount == 106) {
        for (let i = 77; i <= 94; i++) {
            form2Part4Data.push(form2Worksheet.getRow(i).values.slice(3, jumlahPertandinganForm2 * 2 + 4))
        }
    }

    // Form 3

    form3Part4Data = []

    if (form3Worksheet.actualRowCount == 91) {
        for (let i = 67; i <= 81; i++) {
            form3Part4Data.push(form3Worksheet.getRow(i).values.slice(3, jumlahPertandinganForm3 * 2 + 4))
        }
    } else if (form3Worksheet.actualRowCount == 106) {
        for (let i = 77; i <= 94; i++) {
            form3Part4Data.push(form3Worksheet.getRow(i).values.slice(3, jumlahPertandinganForm3 * 2 + 4))
        }
    }

    // Form 4

    form4Part4Data = []

    if (form4Worksheet.actualRowCount == 91) {
        for (let i = 67; i <= 81; i++) {
            form4Part4Data.push(form4Worksheet.getRow(i).values.slice(3, jumlahPertandinganForm4 * 2 + 4))
        }
    } else if (form4Worksheet.actualRowCount == 106) {
        for (let i = 77; i <= 94; i++) {
            form4Part4Data.push(form4Worksheet.getRow(i).values.slice(3, jumlahPertandinganForm4 * 2 + 4))
        }
    }

    // Form 5

    form5Part4Data = []

    if (form5Worksheet.actualRowCount == 91) {
        for (let i = 67; i <= 81; i++) {
            form5Part4Data.push(form5Worksheet.getRow(i).values.slice(3, jumlahPertandinganForm5 * 2 + 4))
        }
    } else if (form5Worksheet.actualRowCount == 106) {
        for (let i = 77; i <= 94; i++) {
            form5Part4Data.push(form5Worksheet.getRow(i).values.slice(3, jumlahPertandinganForm5 * 2 + 4))
        }
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

    // Final Part: Totaled

    startingMerit = 500

    // For each form, if a form has 91 rows, then we'll need to get data from rows 87 to 101
    // For each form, if a form has 106 rows, then we'll need to get data from rows 100 to 117
    // For each row, get the values for column 4

    // Store the data in an array 

    // Form 1

    form1Part5Data = []

    if (form1Worksheet.actualRowCount == 91) {
        for (let i = 87; i <= 101; i++) {
            form1Part5Data.push(form1Worksheet.getRow(i).values[4])
        }
    } else if (form1Worksheet.actualRowCount == 106) {
        for (let i = 100; i <= 117; i++) {
            form1Part5Data.push(form1Worksheet.getRow(i).values[4])
        }
    }

    // Form 2

    form2Part5Data = []

    if (form2Worksheet.actualRowCount == 91) {
        for (let i = 87; i <= 101; i++) {
            form2Part5Data.push(form2Worksheet.getRow(i).values[4])
        }
    } else if (form2Worksheet.actualRowCount == 106) {
        for (let i = 100; i <= 117; i++) {
            form2Part5Data.push(form2Worksheet.getRow(i).values[4])
        }
    }

    // Form 3

    form3Part5Data = []

    if (form3Worksheet.actualRowCount == 91) {
        for (let i = 87; i <= 101; i++) {
            form3Part5Data.push(form3Worksheet.getRow(i).values[4])
        }
    } else if (form3Worksheet.actualRowCount == 106) {
        for (let i = 100; i <= 117; i++) {
            form3Part5Data.push(form3Worksheet.getRow(i).values[4])
        }
    }

    // Form 4

    form4Part5Data = []

    if (form4Worksheet.actualRowCount == 91) {
        for (let i = 87; i <= 101; i++) {
            form4Part5Data.push(form4Worksheet.getRow(i).values[4])
        }
    } else if (form4Worksheet.actualRowCount == 106) {
        for (let i = 100; i <= 117; i++) {
            form4Part5Data.push(form4Worksheet.getRow(i).values[4])
        }
    }

    // Form 5

    form5Part5Data = []

    if (form5Worksheet.actualRowCount == 91) {
        for (let i = 87; i <= 101; i++) {
            form5Part5Data.push(form5Worksheet.getRow(i).values[4])
        }
    } else if (form5Worksheet.actualRowCount == 106) {
        for (let i = 100; i <= 117; i++) {
            form5Part5Data.push(form5Worksheet.getRow(i).values[4])
        }
    }

    // Logging the data

    // console.log("Form 1 Part 5 Data:")
    // console.log(form1Part5Data)
    // console.log("Form 2 Part 5 Data:")
    // console.log(form2Part5Data)
    // console.log("Form 3 Part 5 Data:")
    // console.log(form3Part5Data)
    // console.log("Form 4 Part 5 Data:")
    // console.log(form4Part5Data)
    // console.log("Form 5 Part 5 Data:")
    // console.log(form5Part5Data)

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
            },
            "part5": {
                "data": form1Part5Data
            },

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
            },
            "part5": {
                "data": form2Part5Data
            },
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
            },
            "part5": {
                "data": form3Part5Data
            },
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
            },
            "part5": {
                "data": form4Part5Data
            },
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
            },
            "part5": {
                "data": form5Part5Data
            },

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
