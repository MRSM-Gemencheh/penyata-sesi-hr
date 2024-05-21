import ExcelJS from 'exceljs';

try {
    if (ExcelJS) {
        console.info("ExcelJS loaded successfully!")
    }
} catch (err) {
    console.error("ExcelJS failed to load")
}

console.log("Bundle file has been successfully loaded!")