//Updates the individual team sheets with the stats from
//the offensive and defensive stat sheets
const Excel = require('exceljs');

const workbook = new Excel.Workbook();

workbook.xlsx.readFile('2018CFPickem.xlsx')
.then(() => {
    // function variables
    let i = 2;
    let team = '';
    // function constants
    const offSheet = workbook.getWorksheet('Offense');
    // Loop through teams on the sheet
    while (team != 'N') {
        // Set row
        row01 = offSheet.getRow(i);
        // Get team stat information
        team = row01.getCell(2).value;
        if (team != 'N') {
            let opts = row01.getCell(3).value;
            let otyds = row01.getCell(4).value;
            let opyds = row01.getCell(5).value;
            let oryds = row01.getCell(6).value;
            try {
                let teamSheet = workbook.getWorksheet(team);
                let row02 = teamSheet.getRow(5);
                let row05 = teamSheet.getRow(6);
                //Copy team stats to another row to determine single game stats
                row05.getCell(1).value = row02.getCell(1).value;
                row05.getCell(2).value = row02.getCell(2).value;
                row05.getCell(4).value = row02.getCell(4).value;
                row05.getCell(6).value = row02.getCell(6).value;
                row05.getCell(8).value = row02.getCell(8).value;
                row05.getCell(10).value = row02.getCell(10).value;
                row05.getCell(12).value = row02.getCell(12).value;
                row05.getCell(14).value = row02.getCell(14).value;
                row05.getCell(16).value = row02.getCell(16).value;
                if (otyds != row02.getCell(4).value) {
                    let games = row02.getCell(1).value + 1;
                    row02.getCell(1).value = games;
                    row02.getCell(2).value = opts;
                    row02.getCell(4).value = otyds;
                    row02.getCell(6).value = opyds;
                    row02.getCell(8).value = oryds;
                    row02.commit();
                }
            } catch(err) {
                console.log(err);
                console.log(team);
            }
        }
        // Increment i for row number
        i++;
    }
    workbook.xlsx.writeFile('2018CFPickem.xlsx');
}).then(() => {
    // function variables
    let j = 2;
    let dteam = '';
    // function constants
    const defSheet = workbook.getWorksheet('Defense');
    // Loop through teams on the sheet
    while (dteam != 'N') {
        // Set row
        row03 = defSheet.getRow(j);
        // Get team stat information
        dteam = row03.getCell(2).value;
        if (dteam != 'N') {
            let dpts = row03.getCell(3).value;
            let dtyds = row03.getCell(4).value;
            let dpyds = row03.getCell(5).value;
            let dryds = row03.getCell(6).value;
            try {
                let dteamSheet = workbook.getWorksheet(dteam);
                let row04 = dteamSheet.getRow(5);
                if (dtyds != row04.getCell(4).value) {
                    row04.getCell(10).value = dpts;
                    row04.getCell(12).value = dtyds;
                    row04.getCell(14).value = dpyds;
                    row04.getCell(16).value = dryds;
                    row04.commit();
                }
            } catch(err) {
                console.log(err);
                console.log(dteam);
            }
        }
        // Increment j for row number
        j++;
    }
    workbook.xlsx.writeFile('2018CFPickem.xlsx');
});