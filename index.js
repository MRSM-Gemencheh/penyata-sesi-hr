const ExcelJS = require('exceljs');
const prompt = require('prompt-sync')({sigint: true});
const fs = require('fs');

const workbook = new ExcelJS.Workbook();

// Ask the user for the filename of the Excel file to read from 

year = prompt("Masukkan tahun: (20XX) ")

filename = 'src/Data_Merit_Demerit_HR_' + String(year) + '.xlsx'

async function readExcelFileFromSystem(filename) {

    try {
        await workbook.xlsx.readFile(filename);
    } catch (err) {
        console.log(err)
        console.log("Fail tidak dijumpai! Sila pastikan fail tersebut berada di dalam folder src/ dan nama fail tersebut adalah Data_Merit_Demerit_HR_20XX.xlsx")
        process.exit()
    }

    console.info("Fail berjaya dibaca!")
    readDataFromFile()

}

readExcelFileFromSystem(filename)

async function readDataFromFile() {
    form1Worksheet = workbook.getWorksheet('TING.1')
    form2Worksheet = workbook.getWorksheet('TING.2')
    form3Worksheet = workbook.getWorksheet('TING.3')
    form4Worksheet = workbook.getWorksheet('TING.4')
    form5Worksheet = workbook.getWorksheet('TING.5')

    console.info("Sedang mula membaca data dari fail...")

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
        form1Part3Data.push(form1Worksheet.getRow(i).values.slice(3, 15))
        form2Part3Data.push(form2Worksheet.getRow(i).values.slice(3, 15))
        form3Part3Data.push(form3Worksheet.getRow(i).values.slice(3, 15))
        form4Part3Data.push(form4Worksheet.getRow(i).values.slice(3, 15))
        form5Part3Data.push(form5Worksheet.getRow(i).values.slice(3, 15))
    }

    // Part 4

    // Form 1

    let form1PertandinganNames, form2PertandinganNames, form3PertandinganNames, form4PertandinganNames, form5PertandinganNames
    form1PertandinganNames = []
    form2PertandinganNames = []
    form3PertandinganNames = []
    form4PertandinganNames = []
    form5PertandinganNames = []

    for (let i = 3; i <= 23; i++) {
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
        form1Part4Data.push(form1Worksheet.getRow(i).values.slice(3, 24))
        form2Part4Data.push(form2Worksheet.getRow(i).values.slice(3, 23))
        form3Part4Data.push(form3Worksheet.getRow(i).values.slice(3, 25))
        form4Part4Data.push(form4Worksheet.getRow(i).values.slice(3, 24))
        form5Part4Data.push(form5Worksheet.getRow(i).values.slice(3, 24))
    }

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

    fs.writeFile(`data_json/merit_demerit_data_${year}.json`, JSON.stringify(data), (err) => {
        if (err) {
            console.log(err)
        } else {
            console.log("Data daripada fail Excel berjaya disimpan ke dalam fail JSON!")
            console.log("Fail disimpan di " + `data_json/merit_demerit_data_${year}.json`)
            console.log("Sila taip 'node write.js' untuk mula menjana penyata akhir HR.")
        }
    }
    )

}
