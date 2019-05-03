import * as Excel from 'exceljs'; //bug in types
// const Excel = require('exceljs')

// local variables, for holding data
const file1: string[][] = [];
const file1Index: string[] = [];
const file2: string[][] = [];
const file2Index: any[] = [];
const errors: any[] = [];
const doubleID = process.argv[2] === 'yes' ? true : false;
const file1Name = process.argv[3] || 'file1';
const file2Name = process.argv[4] || 'file2';
const file1Column = `${file1Name}-data`;
const file2Column = `${file2Name}-data`;

if (doubleID) {
    console.log('combining column 1 and 2 into ID')
}

// read excel
const readfile = async (file_XLSX: string, intoArray: any[], index: string[]) => {
    console.log('reading file:', file_XLSX)
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(file_XLSX)
    let worksheet;
    let i = workbook.eachSheet((sheet) => { // loop and find worksheet
        if (sheet && !worksheet) {
            worksheet = sheet;
        }
    });
    if (!worksheet) {
        console.log('No worksheet found');
    }
    worksheet.eachRow(function (row: any) {
        let values = [];
        row.values.forEach((col) => {
            if (typeof col === 'string') {
                values.push(col)
            } else {
                if (col) {
                    values.push(col.text);
                }
            }
        })
        intoArray.push(values)
        if (values.length > 1) {
            if (doubleID) {
                index.push(values[0] + ';' + values[1]);
            } else {
                index.push(values[0]);
            }

        }
    });
}

// compare data and add to error file
const compareFiles = async () => {


    console.log(`Looping ${file1Name}.xlsx`)
    file1Index.forEach((id: string, index: number) => {
        //skip first- this is the header
        if (index % 1000 === 0) { //progress...
            console.log(`Looping ${file1Name}.xlsx - at index:`, index);
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
                            change: file1[0][i],
                            [`${file1Column}`]: rowdata1[i],
                            [`${file2Column}`]: rowdata2[i]
                        });
                    }
                }

            } else {

                if (doubleID) {
                    let vals = id.split(';');
                    errors.push({
                        id1: vals[0],
                        id2: vals[1],
                        change: `In ${file1Name}.xlsx only`,
                        [`${file1Column}`]: 'NA',
                        [`${file2Column}`]: 'NA'
                    });
                } else {
                    errors.push({
                        id: id,
                        change: `In ${file1Name}.xlsx only`,
                        [`${file1Column}`]: 'NA',
                        [`${file2Column}`]: 'NA'
                    });
                }


            }
        }
    });

    console.log(`Looping ${file2Name}.xlsx`)
    file2Index.forEach((id: string, index: number) => {
        //skip first- this is the header
        if (index % 1000 === 0) {
            console.log(`Looping ${file2Name}.xlsx - at index:`, index);
        }
        if (index > 0) {
            let file1row = file1Index.indexOf(id);
            if (file1row === -1) {
                if (doubleID) {
                    let vals = id.split(';');
                    errors.push({
                        id1: vals[0],
                        id2: vals[1],
                        change: `In ${file2Name}.xlsx only`,
                        [`${file1Column}`]: 'NA',
                        [`${file2Column}`]: 'NA'
                    });
                } else {
                    errors.push({
                        id: id,
                        change: `In ${file2Name}.xlsx only`,
                        [`${file1Column}`]: 'NA',
                        [`${file2Column}`]: 'NA'
                    });
                }
            }
        }
    });
}


const generateErrorFile = async () => {
    console.log('generating errorReport')
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
        await readfile(`./${file1Name}.xlsx`, file1, file1Index);
        await readfile(`./${file2Name}.xlsx`, file2, file2Index);
        await compareFiles();
        await generateErrorFile();
        console.log("done")
    } catch (e) {
        console.log("Something failed", e);
    }

}

main();



