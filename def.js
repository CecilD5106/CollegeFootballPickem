//Get the defensive stats for all of the teams in 
//the BCS division in college football
const request = require('request');
const cheerio = require('cheerio');
const Excel = require('exceljs');

let workbook = new Excel.Workbook();

workbook.xlsx.readFile('2018CFPickem.xlsx')
.then(() => {
    request('http://www.espn.com/college-football/statistics/team/_/stat/total/position/defense/sort/totalYards', 
        (error, response, html) => {
            console.log('Begin get defensive stats')
            if (!error && response.statusCode == 200) {
                let oldRank = 0;
                const $ = cheerio.load(html);

                const rows = $('.tablehead').children().children().length;
                let j = 2;
                const worksheet = workbook.getWorksheet('Defense');

                for (let i = 0; i < rows; i++) {
                    let rank = $('.tablehead').children().children().eq(i).children().eq(0).text();
                    const htmlRank = $('.tablehead').children().children().eq(i).children().eq(0).html();
                    if (htmlRank == '&#xA0;') {
                        rank = oldRank;
                    }
                    if (rank != 'RK') {
                        const team = $('.tablehead').children().children().eq(i).children().eq(1).text();
                        const yds = $('.tablehead').children().children().eq(i).children().eq(2).text();
                        const pass = $('.tablehead').children().children().eq(i).children().eq(4).text();
                        const rush = $('.tablehead').children().children().eq(i).children().eq(6).text();
                        const pts = $('.tablehead').children().children().eq(i).children().eq(8).text();
                        let row01 = worksheet.getRow(j);
                        row01.getCell(2).value = team;
                        row01.getCell(3).value = parseInt(pts);
                        row01.getCell(4).value = parseInt(yds);
                        row01.getCell(5).value = parseInt(pass);
                        row01.getCell(6).value = parseInt(rush);
                        j = j + 1;
                    }
                    oldRank = rank;
                }
                workbook.xlsx.writeFile('2018CFPickem.xlsx');
                console.log('End get defensive stats');
            }
        });
});