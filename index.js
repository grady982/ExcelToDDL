const fs = require('fs');
const xlsx = require('xlsx');

// Read the Excel file
const excelFilePath = `C:\\Users\\grady.liu\\Downloads\\columnType.xlsx`;
const workbook = xlsx.readFile(excelFilePath);
console.log(workbook.SheetNames);
const sheetName = workbook.SheetNames[1];
const worksheet = workbook.Sheets[sheetName];

// Convert Excel data to JSON
const excelData = xlsx.utils.sheet_to_json(worksheet);

// Group MSSQL DDL statements by table name
const ddlStatementsByTable = {};

// Create MSSQL DDL statements
excelData.forEach((row) => {
    const isPrimaryKey = row['是否為PK'] === 'Y';
    const isAllowNull = row['允許空值'] === 'Y';
    const allowNull = row['允許空值'] === 'Y' ? 'NULL' : 'NOT NULL';

    let dataType = row['資料型態'];

    // switch(row['資料型態'].toLowerCase()) {
    //     case 'nvarchar(-1)':
    //         dataType = 'nvarchar(max)';
    //         break;
    //     case 'nvarchar(0)':
    //         dataType = 'nvarchar';
    //         break;
    //     case 'varrchar(50)':
    //         dataType = 'varchar(50)';
    //         break;
    //     case 'varchar(-1)':
    //         dataType = 'varchar(max)';
    //         break;
    //     case 'smallint(5,0)':
    //         dataType = 'smallint';
    //         break;
    //     case 'bigint(19,0)':
    //         dataType = 'bigint';
    //         break;
    //     case 'tinyint(3,0)':
    //         dataType = 'tinyint';
    //         break;
    //     case 'int(4)':
    //         dataType = 'int';
    //         break;
    //     case 'datetime2(7)':
    //         dataType = '[datetime2](7)';
    //         break;
    //     case 'datetime(8)':
    //         dataType = 'datetime';
    //         break;
    //     case 'date(3)':
    //         dataType = 'date';
    //         break;
    //     case 'money(19,4)':
    //         dataType = 'money';
    //         break;
    //     case 'flaot':
    //         dataType = 'float';
    //         break;
    //     case 'foat':
    //         dataType = 'float';
    //         break;
    //     case 'datechar(1)':
    //         dataType = 'varchar(1)';
    //         break;
    //     case 'numeric(9,':
    //         dataType = 'numeric(9,0)';
    //         break;
    //     case 'bit(1)':
    //         dataType = 'bit';
    //         break;
    //     case 'bigint(8)':
    //         dataType = 'bigint';
    //         break;
    //     case 'uniqueidentifier(16)':
    //         dataType = 'uniqueidentifier';
    //         break;
    //     default:
    //         dataType = row['資料型態'].toLowerCase();
    //         break;
    // }

    const tableName = row['資料表名稱'].toLowerCase();
    const colName = row['欄位名稱'].toLowerCase();

    let ddlStatement = `    [${colName}] ${dataType} ${allowNull},\n`;

    if (!ddlStatementsByTable[tableName]) {
        ddlStatementsByTable[tableName] = { 
            columns: [], 
            columnStatements: [],
            primaryKeys: [] 
        };
        ddlStatementsByTable[tableName].columnStatements.push(`CREATE TABLE ${row['結構名稱']}.[${tableName}] (\n`);
    }

    const notExist = ddlStatementsByTable[tableName].columns.findIndex(col => col == colName) == -1;

    if (notExist) {
        ddlStatementsByTable[tableName].columnStatements.push(ddlStatement);
        ddlStatementsByTable[tableName].columns.push(colName);

        if (isPrimaryKey && !isAllowNull && dataType != 'varchar(max)' && dataType != 'nvarchar(max)') {
            if (ddlStatementsByTable[tableName].primaryKeys.length < 32) {
                ddlStatementsByTable[tableName].primaryKeys.push(`[${colName}],`);    
            }
        }
    }

});

// Save all MSSQL DDL statements to a single text file
const outputFilePath = 'C:\\Users\\grady.liu\\Desktop\\Testddl.sql';
fs.writeFileSync(outputFilePath, '');

Object.entries(ddlStatementsByTable).forEach(([tableName, statements]) => {
    const primaryKeyStatement = statements.primaryKeys.length > 0
        ? `    CONSTRAINT PK_${tableName} PRIMARY KEY (${statements.primaryKeys.join('').slice(0, -1)})\n`
        : '';

    fs.writeFileSync(outputFilePath, statements.columnStatements.concat(primaryKeyStatement, ');\n\n').join(''), { flag: 'a' });
});


