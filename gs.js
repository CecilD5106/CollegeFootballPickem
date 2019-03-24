//Updates the individual team sheets with the game stats from
//the team stat sheets
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
    row01 = offSheet.getRow(i);
    team = row01.getCell(2).value;
    while (team != 'N') {
        let j = 12;
        let teamSheet = workbook.getWorksheet(team);
        let row02 = teamSheet.getRow(7);
        let row03 = teamSheet.getRow(j);
        let cell02 = row03.getCell(1).value;
        while (cell02 != null) {
            j++;
            row03 = teamSheet.getRow(j);
            cell02 = row03.getCell(1).value;
        }
        //Determine if the team had a bye week
        if (row02.getCell(4).value.result != 0) {
            // Put game stats in the next empty row after row 12
            row03.getCell(1).value = row02.getCell(2).value.result;
            // Get points for value
            let pf = row03.getCell(1).value;
            row03.getCell(2).value = row02.getCell(4).value.result;
            row03.getCell(3).value = row02.getCell(6).value.result;
            row03.getCell(4).value = row02.getCell(8).value.result;
            row03.getCell(5).value = row02.getCell(10).value.result;
            //Get points against value
            let pa = row03.getCell(5).value;
            row03.getCell(6).value = row02.getCell(12).value.result;
            row03.getCell(7).value = row02.getCell(14).value.result;
            row03.getCell(8).value = row02.getCell(16).value.result;
            // Compare points for and points against to determine the winner
            // and add points to the win column
            if (pf > pa) {
                row03.getCell(9).value = 1;
            } else {
                row03.getCell(9).value = 0;
            }
        }
        //Increment i
        i++;
        //Update row01
        row01 = offSheet.getRow(i);
        //Update team
        team = row01.getCell(2).value;
    }
    workbook.xlsx.writeFile('2018CFPickem.xlsx');
});