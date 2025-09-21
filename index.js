import fs from 'fs';
import path from 'path';
import * as XLSX from 'xlsx'
import iconv from 'iconv-lite';
import {parse} from 'csv-parse/sync';

const inputFile = process.argv[2];
if (!inputFile) {
    console.error('Usage: node script210925.bundle.js 测试.csv');
    process.exit(1);
}

const buffer = fs.readFileSync(inputFile);
const csvData = iconv.decode(buffer, 'gb18030');

const records = parse(csvData, {
    columns: true,
    skip_empty_lines: true,
});

const excelRows = [];
excelRows.push(['名称', '经度', '纬度', '海拔', '文本显示风格', '图标样式', 'name', 'sex', 'nation', 'id', 'address', 'relation', 'type', 'change_time']);

for (const rec of records) {
    const comment = commentToJson(rec.Comment);
    const users = extractUsers(comment);
    users.forEach(user => {
        excelRows.push([
            rec['名称'],
            rec['经度'],
            rec['纬度'],
            rec['海拔'],
            rec['文本显示风格'],
            rec['图标样式'],
            user.user,
            user.sex,
            user.nation,
            user.id,
            user.address,
            user.householder_re,
            user.type,
            user.change_time
        ]);
    });
}

const newWs = XLSX.utils.aoa_to_sheet(excelRows);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, newWs, 'Processed');

const outputFile =
    path.join(
        path.dirname(inputFile),
        path.basename(inputFile, path.extname(inputFile)) + '.xlsx'
    );

XLSX.writeFile(wb, outputFile);

console.log(`Processed file written to ${outputFile}`);


function commentToJson(comment) {
    return JSON.parse(comment.replace('<?ovital_ct name="测试">', ''));
}

function extractUsers(data) {
    const indices = new Set();
    for (const key in data) {
        const match = key.match(/\d+$/);
        if (match) indices.add(match[0]);
    }
    return Array.from(indices).map(index => ({
        user: data[`user${index}`],
        sex: data[`sex${index}`],
        nation: data[`nation${index}`],
        id: data[`id${index}`],
        address: data[`address${index}`],
        householder_re: data[`householder_re${index}`],
        type: data[`type${index}`],
        change_time: data[`change_time${index}`],
    }));
}