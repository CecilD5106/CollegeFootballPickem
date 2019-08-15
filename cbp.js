//Updates the probabilities based on the stats in the 
//college pickem worksheet for the week
const Excel = require('exceljs');

const workbook = new Excel.Workbook();

workbook.xlsx.readFile('2018CFPickem.xlsx')
.then(() => {
    let week = '';
    //Get the name of the current pick worksheet
    const weekSheet = workbook.getWorksheet('CurPick');
    let row01 = weekSheet.getRow(1);
    week = row01.getCell(1).value;
    //console.log(week);
    const cpSheet = workbook.getWorksheet(week);
    for (let j = 3; j < 26; j = j + 2) {
        //Get Predictors worksheet
        const pSheet = workbook.getWorksheet('Predictors');
        //Get Predictors
        let pRow = pSheet.getRow(j);
        let pHigh = pRow.getCell(19).value.result;
        let pLow = 1 - pHigh;
        for (let k = 5; k < 124; k = k + 3) {
            let vRow = cpSheet.getRow(k);
            let hRow = cpSheet.getRow(k + 1);
            let vStat = 0;
            let hStat = 0;
            switch (j) {
                //Total offensive yards
                case 3:
                    vStat = vRow.getCell(7).value.result;
                    hStat = hRow.getCell(7).value.result;
                    if (vStat > hStat) {
                        vRow.getCell(29).value = pHigh;
                        hRow.getCell(29).value = pLow;
                    } else {
                        vRow.getCell(29).value = pLow;
                        hRow.getCell(29).value = pHigh;
                    }
                    vRow.commit();
                    hRow.commit();
                    break;
                //Total offensive passing yards
                case 5:
                    vStat = vRow.getCell(9).value.result;
                    hStat = hRow.getCell(9).value.result;
                    if (vStat > hStat) {
                        vRow.getCell(30).value = pHigh;
                        hRow.getCell(30).value = pLow;
                    } else {
                        vRow.getCell(30).value = pLow;
                        hRow.getCell(30).value = pHigh;
                    }
                    vRow.commit();
                    hRow.commit();
                    break;
                //Total offensive rushing yards
                case 7:
                    vStat = vRow.getCell(11).value.result;
                    hStat = hRow.getCell(11).value.result;
                    if (vStat > hStat) {
                        vRow.getCell(31).value = pHigh;
                        hRow.getCell(31).value = pLow;
                    } else {
                        vRow.getCell(31).value = pLow;
                        hRow.getCell(31).value = pHigh;
                    }
                    vRow.commit();
                    hRow.commit();
                    break;
                //Offensive points per yard
                case 9:
                    vStat = vRow.getCell(5).value.result;
                    hStat = hRow.getCell(5).value.result;
                    if (vStat > hStat) {
                        vRow.getCell(32).value = pHigh;
                        hRow.getCell(32).value = pLow;
                    } else {
                        vRow.getCell(32).value = pLow;
                        hRow.getCell(32).value = pHigh;
                    }
                    vRow.commit();
                    hRow.commit();
                    break;
                //Defensive total yards allowed
                case 11:
                    vStat = vRow.getCell(16).value.result;
                    hStat = hRow.getCell(16).value.result;
                    if (vStat < hStat) {
                        vRow.getCell(33).value = pHigh;
                        hRow.getCell(33).value = pLow;
                    } else {
                        vRow.getCell(33).value = pLow;
                        hRow.getCell(33).value = pHigh;
                    }
                    vRow.commit();
                    hRow.commit();
                    break;
                //Defensive passing yards allowed
                case 13:
                    vStat = vRow.getCell(18).value.result;
                    hStat = hRow.getCell(18).value.result;
                    if (vStat < hStat) {
                        vRow.getCell(34).value = pHigh;
                        hRow.getCell(34).value = pLow;
                    } else {
                        vRow.getCell(34).value = pLow;
                        hRow.getCell(34).value = pHigh;
                    }
                    vRow.commit();
                    hRow.commit();
                    break;
                //Defensive rushing yards allowed
                case 15:
                    vStat = vRow.getCell(20).value.result;
                    hStat = hRow.getCell(20).value.result;
                    if (vStat < hStat) {
                        vRow.getCell(35).value = pHigh;
                        hRow.getCell(35).value = pLow;
                    } else if (vStat > hStat) {
                        vRow.getCell(35).value = pLow;
                        hRow.getCell(35).value = pHigh;
                    } else {
                        vRow.getCell(35).value = 0;
                        hRow.getCell(35).value = 0;
                    }
                    vRow.commit();
                    hRow.commit();
                    break;
                //Defesive points per yard
                case 17:
                    vStat = vRow.getCell(14).value.result;
                    hStat = hRow.getCell(14).value.result;
                    if (vStat < hStat) {
                        vRow.getCell(36).value = pHigh;
                        hRow.getCell(36).value = pLow;
                    } else if (vStat > hStat) {
                        vRow.getCell(36).value = pLow;
                        hRow.getCell(36).value = pHigh;
                    } else {
                        vRow.getCell(36).value = 0;
                        hRow.getCell(36).value = 0;
                    }
                    vRow.commit();
                    hRow.commit();
                    break;
                //Ranking
                case 19:
                    vStat = vRow.getCell(21).value;
                    hStat = hRow.getCell(21).value;
                    if (vStat < hStat) {
                        vRow.getCell(37).value = pHigh;
                        hRow.getCell(37).value = pLow;
                    } else if (vStat > hStat) {
                        vRow.getCell(37).value = pLow;
                        hRow.getCell(37).value = pHigh;
                    } else {
                        vRow.getCell(37).value = 0;
                        hRow.getCell(37).value = 0;
                    }
                    vRow.commit();
                    hRow.commit();
                    break;
                //Home team
                case 21:
                    vStat = vRow.getCell(22).value;
                    hStat = hRow.getCell(22).value;
                    if (vStat > hStat) {
                        vRow.getCell(38).value = pHigh;
                        hRow.getCell(38).value = pLow;
                    } else if (vStat < hStat) {
                        vRow.getCell(38).value = pLow;
                        hRow.getCell(38).value = pHigh;
                    } else {
                        vRow.getCell(37).value = 0;
                        hRow.getCell(37).value = 0;
                    }
                    vRow.commit();
                    hRow.commit();
                    break;
                //Last three game record
                case 23:
                    vStat = vRow.getCell(23).value;
                    hStat = hRow.getCell(23).value;
                    if (vStat > hStat) {
                        vRow.getCell(39).value = pHigh;
                        hRow.getCell(39).value = pLow;
                    } else if (vStat < hStat) {
                        vRow.getCell(39).value = pLow;
                        hRow.getCell(39).value = pHigh;
                    } else {
                        vRow.getCell(39).value = 0;
                        hRow.getCell(39).value = 0;
                    }
                    vRow.commit();
                    hRow.commit();
                    break;
                //Last five game record
                case 25:
                    vStat = vRow.getCell(24).value;
                    hStat = hRow.getCell(24).value;
                    if (vStat > hStat) {
                        vRow.getCell(40).value = pHigh;
                        hRow.getCell(40).value = pLow;
                    } else if (vStat < hStat) {
                        vRow.getCell(40).value = pLow;
                        hRow.getCell(40).value = pHigh;
                    } else {
                        vRow.getCell(40).value = 0;
                        hRow.getCell(40).value = 0;
                    }
                    vRow.commit();
                    hRow.commit();
                    break;
                default:
                    console.log('Break');
            }
        }
    }
    // Save workbook
    workbook.xlsx.writeFile('2019CFPickem.xlsx');
});
