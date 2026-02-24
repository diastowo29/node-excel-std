const Excel = require('exceljs');
const fs = require('fs');
const readline = require('readline');
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

let daysWorkHourCell = 'F11'
let mealAllowanceCell = 'F12'
let transportAllowanceCell = 'F13'
// let attendanceRewardCell = 'F14'
// let ontimeAllowanceCell = 'F15'

/* NEW ON V2 */
let productivityBonusCell = 'F14'
let awayAllowanceCell = 'F15'
let mealAllowanceWeekendCell = 'F16'
let transportAllowanceWeekendCell = 'F17'
let productivityBonusWeekendCell = 'F18'
let totalOvertimeWeekday = 'F19'
/* NEW ON V2 */

// let overtimeWeekdayCell = 'F17'
// let overtimeSaturdayCell = 'F18'
// let overtimeSundayCell = 'F19'
// let minustimepermonthCell = 'F20'

let publicHolidayCell = 'R11'
let saturdaySundayCell = 'R12'
let annualLeaveCell = 'R13'
let compassionateLeaveCell = 'R14'
let paidSickCell = 'R15'
let daysDeductedCell = 'R16'
let totalDaysCell = 'R17'
let availableWeekdaysCell = 'R18'
let overtimeWeekdayUCell = 'R19'
let overtimeSatSunCell = 'R20'
let totalWorkDaysCell = 'R21'

// let productivityCell = 'T67'
// let hseCell = 'U67'


let nameCell = 'C5'
let empNoCell = 'C6'
let folderPrefix = 'D:/WORK/NodeJS/nodejs-excel-std/ISC_PAYROLL';

const sheetName = process.argv[2];
const year = process.argv[3];
const excelMonthFolder = process.argv[4];
const sheetFallbackName = process.argv[2];
readFileFolder();

/* rl.question('Sheet name ? ', function (sheetInput) {
    rl.question('Sheet alternative name ? ', function (sheetAltInput) {
        rl.question('What year ? ', function (yearInput) {
            rl.question('Excel month folder name? ', function(folderInput) {
                sheetName = sheetInput;
                sheetFallbackName = sheetAltInput;
                excelMonthFolder = folderInput;
                year = yearInput;
                // console.log('start')
                // console.log(sheetName)
                // console.log(sheetFallbackName)
                // console.log(excelMonthFolder)
                readFileFolder();
                rl.close();
            });
        });
    });
}); */


// const readline = require("readline");
// const rl = readline.createInterface({
//     input: process.stdin,
//     output: process.stdout
// });

// rl.question("Sheet name? ", function(sheetNameInput) {
// 	rl.question('Sheet Fallback name? ', function(sheetFallbackNameInput) {
// 	    rl.question("Month Folder? ", function(monthFolderInput) {
// 	        sheetName = sheetNameInput;
// 	        sheetFallbackName = sheetFallbackNameInput;
// 	        excelMonthFolder = monthFolderInput;
// 	        rl.close();
// 	    });
// 	})
// });

// rl.on('close', (input) => {
// 	console.log('sheetName: %s', sheetName);
// 	console.log('sheetFallbackName: %s', sheetFallbackName)
// 	console.log('excelMonthFolder: %s', excelMonthFolder);
//     readFileFolder();
// })

let newColumn = {};
let newColumnArray = [];

let newColumnsHeader = [
    {header: 'Employee Number', key: 'empNoCell', width: 10},
    {header: 'Name', key: 'nameCell', width: 10},
    {header: 'Days Worked / Hours', key: 'daysWorkHourCell', width: 10},
    {header: 'Meal Allowance', key: 'mealAllowanceCell', width: 10},
    {header: 'Transport Allowance', key: 'transportAllowanceCell', width: 10},

    {header: 'Productivity Bonus', key: 'productivityBonusCell', width: 10},
    {header: 'Away Allowance', key: 'awayAllowanceCell', width: 10},
    {header: 'Meals Allowance (Weekend, PH)', key: 'mealAllowanceWeekendCell', width: 10},
    {header: 'Transport Allowance (Weekend, PH)', key: 'transportAllowanceWeekendCell', width: 10},
    {header: 'Productivity Bonus (Weekday,end,PH)', key: 'productivityBonusWeekendCell', width: 10},
    {header: 'Total Overtime (Weekday,end,PH)', key: 'totalOvertimeWeekday', width: 10},

    {header: 'Public Holiday', key: 'publicHolidayCell', width: 10},
    {header: 'Saturday & Sunday', key: 'saturdaySundayCell', width: 10},
    {header: 'Annual Leave (AL / SAL)', key: 'annualLeaveCell', width: 10},
    {header: 'Compassionate Leave', key: 'compassionateLeaveCell', width: 10},
    {header: 'Paid Sick', key: 'paidSickCell', width: 10},
    {header: 'Days Deducted (UnPaid)', key: 'daysDeductedCell', width: 10},
    {header: 'Total Days Current Month', key: 'totalDaysCell', width: 10},
    {header: 'Available Weekdays', key: 'availableWeekdaysCell', width: 10},
    {header: 'Overtime: Weekday', key: 'overtimeWeekdayUCell', width: 10},
    {header: 'Overtime: Sat, Sun, PH', key: 'overtimeSatSunCell', width: 10},
    {header: 'Total Work Days', key: 'totalWorkDaysCell', width: 10}
]


async function readExcel (filename) {
    var workbook = new Excel.Workbook();
    var worksheet;
    await workbook.xlsx.readFile(filename)
        .then(function() {
            if (workbook.getWorksheet(sheetName) === undefined) {
                worksheet = workbook.getWorksheet(sheetFallbackName);
            } else {
                worksheet = workbook.getWorksheet(sheetName);
            }

            if (!worksheet) {
                console.log('Sheet not found: ' + filename);
                return;
            }

            console.log('reading file: ' + worksheet.getCell(empNoCell).value + ' - ' 
                + worksheet.getCell(nameCell).value + '...');

            newColumn = {
                name: ((worksheet.getCell(nameCell).value.hasOwnProperty('formula')) ? worksheet.getCell(nameCell).value.result || 0 : worksheet.getCell(nameCell).value),
                empno: ((worksheet.getCell(empNoCell).value.hasOwnProperty('formula')) ? worksheet.getCell(empNoCell).value.result || 0 : worksheet.getCell(empNoCell).value),
                dayswork: ((worksheet.getCell(daysWorkHourCell).value.hasOwnProperty('formula')) ? worksheet.getCell(daysWorkHourCell).value.result || 0 : worksheet.getCell(daysWorkHourCell).value),
                meal: ((worksheet.getCell(mealAllowanceCell).value.hasOwnProperty('formula')) ? worksheet.getCell(mealAllowanceCell).value.result || 0 : worksheet.getCell(mealAllowanceCell).value),
                transport: ((worksheet.getCell(transportAllowanceCell).value.hasOwnProperty('formula')) ? worksheet.getCell(transportAllowanceCell).value.result || 0 : worksheet.getCell(transportAllowanceCell).value),
                
                productivity: ((worksheet.getCell(productivityBonusCell).value.hasOwnProperty('formula')) ? worksheet.getCell(productivityBonusCell).value.result || 0 : worksheet.getCell(productivityBonusCell).value),
                away: ((worksheet.getCell(awayAllowanceCell).value.hasOwnProperty('formula')) ? worksheet.getCell(awayAllowanceCell).value.result || 0 : worksheet.getCell(awayAllowanceCell).value),
                mealWeekend: ((worksheet.getCell(mealAllowanceWeekendCell).value.hasOwnProperty('formula')) ? worksheet.getCell(mealAllowanceWeekendCell).value.result || 0 : worksheet.getCell(mealAllowanceWeekendCell).value),
                transportWeekend: ((worksheet.getCell(transportAllowanceWeekendCell).value.hasOwnProperty('formula')) ? worksheet.getCell(transportAllowanceWeekendCell).value.result || 0 : worksheet.getCell(transportAllowanceWeekendCell).value),
                productivityWeekend: ((worksheet.getCell(productivityBonusWeekendCell).value.hasOwnProperty('formula')) ? worksheet.getCell(productivityBonusWeekendCell).value.result || 0 : worksheet.getCell(productivityBonusWeekendCell).value),
                totalOvertime: ((worksheet.getCell(totalOvertimeWeekday).value.hasOwnProperty('formula')) ? worksheet.getCell(totalOvertimeWeekday).value.result || 0 : worksheet.getCell(totalOvertimeWeekday).value),

                // productivity: ((worksheet.getCell(productivityCell).value.hasOwnProperty('formula')) || 
                // 	(worksheet.getCell(productivityCell).value.hasOwnProperty('sharedFormula')) ? 
                // 	worksheet.getCell(productivityCell).value.result || 0 : 
                // 	worksheet.getCell(productivityCell).value),
                // hse: ((worksheet.getCell(hseCell).value.hasOwnProperty('formula')) || 
                // 	(worksheet.getCell(hseCell).value.hasOwnProperty('sharedFormula')) ?
                // 	worksheet.getCell(hseCell).value.result || 0 : 
                // 	worksheet.getCell(hseCell).value),
                
                publicHoliday: ((worksheet.getCell(publicHolidayCell).value.hasOwnProperty('formula')) ? worksheet.getCell(publicHolidayCell).value.result || 0 : worksheet.getCell(publicHolidayCell).value),
                saturdaySunday: ((worksheet.getCell(saturdaySundayCell).value.hasOwnProperty('formula')) ? worksheet.getCell(saturdaySundayCell).value.result || 0 : worksheet.getCell(saturdaySundayCell).value),
                annualLeave: ((worksheet.getCell(annualLeaveCell).value.hasOwnProperty('formula')) ? worksheet.getCell(annualLeaveCell).value.result || 0 : worksheet.getCell(annualLeaveCell).value),
                compassionate: ((worksheet.getCell(compassionateLeaveCell).value.hasOwnProperty('formula')) ? worksheet.getCell(compassionateLeaveCell).value.result || 0 : worksheet.getCell(compassionateLeaveCell).value),
                paidSick: ((worksheet.getCell(paidSickCell).value.hasOwnProperty('formula')) ? worksheet.getCell(paidSickCell).value.result || 0 : worksheet.getCell(paidSickCell).value),
                daysDeducted: ((worksheet.getCell(daysDeductedCell).value.hasOwnProperty('formula')) ? worksheet.getCell(daysDeductedCell).value.result || 0 : worksheet.getCell(daysDeductedCell).value),
                totalDays: ((worksheet.getCell(totalDaysCell).value.hasOwnProperty('formula')) ? worksheet.getCell(totalDaysCell).value.result || 0 : worksheet.getCell(totalDaysCell).value),
                availableWeekdays: ((worksheet.getCell(availableWeekdaysCell).value.hasOwnProperty('formula')) ? worksheet.getCell(availableWeekdaysCell).value.result || 0 : worksheet.getCell(availableWeekdaysCell).value),
                overtimeWeekdayU: ((worksheet.getCell(overtimeWeekdayUCell).value.hasOwnProperty('formula')) ? worksheet.getCell(overtimeWeekdayUCell).value.result || 0 : worksheet.getCell(overtimeWeekdayUCell).value),
                overtimeSatSun: ((worksheet.getCell(overtimeSatSunCell).value.hasOwnProperty('formula')) ? worksheet.getCell(overtimeSatSunCell).value.result || 0 : worksheet.getCell(overtimeSatSunCell).value),
                totalWorkDays: ((worksheet.getCell(totalWorkDaysCell).value.hasOwnProperty('formula')) ? worksheet.getCell(totalWorkDaysCell).value.result || 0 : worksheet.getCell(totalWorkDaysCell).value),
            };
        });
}


async function writeExcel(dataArray){

    console.log('===== Generating Summary File =====')
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet(sheetName);

    worksheet.columns = newColumnsHeader;
    
    for (var i=0; i<dataArray.length; i++) {
        // console.log(dataArray[i].name + " : " + dataArray[i].productivity + " : " + dataArray[i].hse)
        worksheet.addRow({
            nameCell: dataArray[i].name,
            empNoCell: dataArray[i].empno,
            daysWorkHourCell: dataArray[i].dayswork,
            mealAllowanceCell: dataArray[i].meal,
            transportAllowanceCell: dataArray[i].transport,
            // attendanceRewardCell: dataArray[i].attendance,
            // ontimeAllowanceCell: dataArray[i].ontime,
            
            productivityBonusCell: dataArray[i].productivity,
            awayAllowanceCell: dataArray[i].away,
            mealAllowanceWeekendCell: dataArray[i].mealWeekend,
            transportAllowanceWeekendCell: dataArray[i].transportWeekend,
            productivityBonusWeekendCell: dataArray[i].productivityWeekend,
            totalOvertimeWeekday: dataArray[i].totalOvertime,
            // productivityCell: dataArray[i].productivity,
            // hseCell: dataArray[i].hse,
            publicHolidayCell: dataArray[i].publicHoliday,
            saturdaySundayCell: dataArray[i].saturdaySunday,
            annualLeaveCell: dataArray[i].annualLeave,
            compassionateLeaveCell: dataArray[i].compassionate,
            paidSickCell: dataArray[i].paidSick,
            daysDeductedCell: dataArray[i].daysDeducted,
            totalDaysCell: dataArray[i].totalDays,
            availableWeekdaysCell: dataArray[i].availableWeekdays,
            overtimeWeekdayUCell: dataArray[i].overtimeWeekdayU,
            overtimeSatSunCell: dataArray[i].overtimeSatSun,
            totalWorkDaysCell: dataArray[i].totalWorkDays
        });
    }

    // save under export.xlsx
    await workbook.xlsx.writeFile(folderPrefix + '/Summary Result/NODEJS-SUMMARY-' + sheetName + '.xlsx');

    console.log("File is written");
};

async function readFileFolder () {
    var fileCounter = 0;
    const excelSource = folderPrefix + '/'+ year +'/' + excelMonthFolder + '/';
    console.log('EXCEL SOURCE FOLDER: %s', excelSource)
    fs.readdir(excelSource, async (err, files) => {
        let newFilesArray = [];
        await files.forEach(async file => {
            if (!file.includes('~')) {
                let new_filename = excelSource + file
                newFilesArray.push(new_filename);
            }
        })
        console.log(newFilesArray.length)
        for (var i=0; i<newFilesArray.length; i++) {
            await readExcel(newFilesArray[i]);
            fileCounter++;
            newColumnArray.push(newColumn);
            if (fileCounter == newFilesArray.length) {
                writeExcel(newColumnArray)
            }
            newColumn = {};
        }
    });
}