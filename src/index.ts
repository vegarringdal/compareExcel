import * as Excel from 'exceljs'; //bug in types
// const Excel = require('exceljs')

// local variables, for holding data
const file1: string[][] = [];
const file1Index: string[] = [];
const file2: string[][] = [];
const file2Index: any[] = [];
const errors: any[] = [];


// read excel
const readfile = async (file_XLSX: string, intoArray: any[], index: string[]) => {
    console.log('reading excel file', file_XLSX)
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(file_XLSX)
    let worksheet = workbook.getWorksheet(1);
    worksheet.eachRow({ includeEmpty: true }, function (row: any) {
        intoArray.push(row.values)
        if (row.values.length > 2) {
            index.push(row.values[1]);
        }
    });
}

// compare data and add to error file
const compareFiles = async () => {
    console.log('comparing file 1 to file 2')
    file1Index.forEach((id: string, index: number) => {
        //skip first- this is the header
        if (index % 1000 === 0) { //progress...
            console.log('comparing file 1 to file 2 - at index', index);
        }
        if (index > 0) {
            let file2row = file2Index.indexOf(id);
            if (file2row !== -1) {
                const rowdata1 = file1[index];
                const rowdata2 = file2[file2row];
                for (let i = 2; i < rowdata1.length; i++) {
                    // skip 2, 0 is always null and next is ID
                    if (rowdata1[i] !== rowdata2[i]) {
                        errors.push({
                            id: id,
                            column: file1[0][i],
                            file1data: rowdata1[i],
                            file2data: rowdata2[i]
                        });
                    }
                }

            } else {
                errors.push({
                    id: id,
                    column: "ONLY IN FILE1",
                    file1data: 'NA',
                    file2data: 'NA'
                });
            }
        }
    });

    console.log('comparing file 2 to file 1')
    file2Index.forEach((id: string, index: number) => {
        //skip first- this is the header
        if (index % 1000 === 0) {
            console.log('comparing file 2 to file 1 - at index', index);
        }
        if (index > 0) {
            let file1row = file1Index.indexOf(id);
            if (file1row === -1) {
                errors.push({
                    id: id,
                    column: "ONLY IN FILE2",
                    file1data: 'NA',
                    file2data: 'NA'
                });
            }
        }
    });
}


const generateErrorFile = async () => {
    console.log('generating excel file')
    if (errors.length > 0) {

        const workbook = new Excel.stream.xlsx.WorkbookWriter({
            filename: `./errorReport-${new Date().getTime()}.xlsx`,
            useStyles: true
        });
        const worksheet = workbook.addWorksheet('errors', {
            views: [
                { state: 'frozen', ySplit: 1 }
            ]
        });

        // generate columns
        const columns = [];
        for (const k in errors[0]) {
            if (errors[0] && errors[0][k] !== undefined) {
                columns.push({
                    header: k,
                    key: k,
                    width: 10,
                    style: {
                        font: { name: 'Calibri Light', size: 10 }
                    }
                });
            }
        }
        worksheet.columns = columns;
        errors.forEach((element: any) => {
            worksheet.addRow(element);
        });

        worksheet.autoFilter = {
            from: {
                row: 1,
                column: 1
            },
            to: {
                row: errors.length,
                column: columns.length
            }
        };

        worksheet.getRow(1).font = { bold: true };

        //style it
        let rowValue = '';
        let toggle = true;
        worksheet.eachRow((row: any, _rowNumber: number) => {

            // toggle per ID (column A)
            let rowValueTemp: string = row.getCell(1) as any;
            if (rowValue !== rowValueTemp) {
                toggle = toggle ? false : true;
            }
            rowValue = rowValueTemp;

            row.eachCell({ includeEmpty: true }, (cell: any, _colNumber: number) => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };

                if (toggle) {
                    cell.fill = <any>{
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFE0E0E0' }
                    };
                }
            });

        });

        await worksheet.commit();
        await workbook.commit();
    } else {
        console.log("skipping errorReport, no errors")
    }


}


const main = async () => {
    try {
        await readfile('./file1.xlsx', file1, file1Index);
        await readfile('./file2.xlsx', file2, file2Index);
        await compareFiles();
        await generateErrorFile();
        console.log("done")
    } catch (e) {
        console.log("Something failed", e);
    }

}

main();



