const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 讀取Excel文件
const chengYuPath = path.join(__dirname, '..', 'docs', '成語.xlsx');
const guShiPath = path.join(__dirname, '..', 'docs', '故事.xlsx');

const output = {
  成語: {},
  故事: {}
};

let chengYuCount = 0;
let guShiCount = 0;

// 處理成語Excel文件
if (fs.existsSync(chengYuPath)) {
  const workbook = XLSX.readFile(chengYuPath);

  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];

    // 讀取資料
    const data = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

    // 轉換資料格式 - 使用陣列存儲
    const sheetData = data
      .filter(row => row.成語 || Object.values(row).some(v => v))
      .map(row => ({
        sheet: sheetName,
        tag: row.tag || '',
        成語: row.成語 || '',
        注音: row.注音 || '',
        解釋: row.解釋 || '',
        造句: row.造句 || ''
      }));

    if (sheetData.length > 0) {
      output.成語[sheetName] = sheetData;
      chengYuCount += sheetData.length;
      console.log(`✓ Sheet: ${sheetName}`);
      console.log(`  - 成語: ${sheetData.length} 條`);
    }
  });
} else {
  console.warn(`⚠ 成語.xlsx 找不到！`);
}

// 處理故事Excel文件
if (fs.existsSync(guShiPath)) {
  const workbook = XLSX.readFile(guShiPath);

  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];

    // 讀取資料
    const data = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

    // 轉換資料格式 - 使用陣列存儲
    const sheetData = data
      .filter(row => row.標題 || Object.values(row).some(v => v))
      .map(row => ({
        sheet: sheetName,
        tag: row.tag || '',
        標題: row.標題 || '',
        故事: row.故事 || ''
      }));

    if (sheetData.length > 0) {
      output.故事[sheetName] = sheetData;
      guShiCount += sheetData.length;
      console.log(`✓ Sheet: ${sheetName}`);
      console.log(`  - 故事: ${sheetData.length} 條`);
    }
  });
} else {
  console.warn(`⚠ 故事.xlsx 找不到！`);
}

// 輸出到文件
const outputPath = path.join(__dirname, '..', 'json', 'chinese.json');
fs.writeFileSync(outputPath, JSON.stringify(output, null, 2), 'utf-8');

console.log(`\n✓ 轉換成功！`);
console.log(`\n統計資訊:`);
console.log(`  - 成語: ${chengYuCount} 條`);
console.log(`  - 故事: ${guShiCount} 條`);
console.log(`  - 合計: ${chengYuCount + guShiCount} 條`);
console.log(`\n輸出文件: ${outputPath}`);
console.log(`\n輸出結構:`);
console.log(`{`);
console.log(`  "成語": {`);
Object.keys(output.成語).forEach(sheetName => {
  console.log(`    "${sheetName}": [${output.成語[sheetName].length} 項]`);
});
console.log(`  },`);
console.log(`  "故事": {`);
Object.keys(output.故事).forEach(sheetName => {
  console.log(`    "${sheetName}": [${output.故事[sheetName].length} 項]`);
});
console.log(`  }`);
console.log(`}`);
