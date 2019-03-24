//Puts the team stats in the college pickem
//worksheet for the week
const Excel = require('exceljs');

const workbook = new Excel.Workbook();

workbook.xlsx.readFile('2018CFPickem.xlsx')
.then(() => {
    let week = '';
    //Get the name of the current pick worksheet
    const weekSheet = workbook.getWorksheet('CurPick');
    let row01 = weekSheet.getRow(1);
    week = row01.getCell(1).value;
    const cfpSheet = workbook.getWorksheet(week);
    //Row variable and team name varible
    let j = 5
    let team = '';
    //Get offense and defense worksheets
    const offSheet = workbook.getWorksheet('Offense');
    const defSheet = workbook.getWorksheet('Defense');
    //Loop through the teams on the current pick worksheet
    while (team != 'X') {
        let row02 = cfpSheet.getRow(j);
        team = row02.getCell(1).value;
        console.log(team);
        if (team != 'N' && team != 'X') {
            let k = 2;
            let offTeam = '';
            //Loop through the teams on the offense worksheet until the team
            //is the same as the team on the current pick worksheet
            while (offTeam != 'N') {
                let row03 = offSheet.getRow(k);
                offTeam = row03.getCell(2).value;
                if (offTeam == team) {
                    row02.getCell(3).value = row03.getCell(3).value;
                    row02.getCell(6).value = row03.getCell(4).value;
                    row02.getCell(8).value = row03.getCell(5).value;
                    row02.getCell(10).value = row03.getCell(6).value;
                    row02.commit();
                }
                k = k + 1;
            }
            let l = 2;
            let defTeam = '';
            //Loop through the teams on the defense worksheet until the team
            //is the same as the team on the current pick worksheet
            while (defTeam != 'N') {
                let row04 = defSheet.getRow(l);
                defTeam = row04.getCell(2).value;
                if (defTeam == team) {
                    row02.getCell(12).value = row04.getCell(3).value;
                    row02.getCell(15).value = row04.getCell(4).value;
                    row02.getCell(17).value = row04.getCell(5).value;
                    row02.getCell(19).value = row04.getCell(6).value;
                    row02.commit();
                }
                l = l + 1;
            }
        }
        j = j + 1;
    }
    workbook.xlsx.writeFile('2018CFPickem.xlsx');
});