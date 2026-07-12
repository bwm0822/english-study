const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 讀取Excel文件
const filePath = path.join(__dirname, 'docs', 'test.xlsx');
const workbook = XLSX.readFile(filePath);

// 取得指定的sheet
const sheetName = 'Abbie-全民英檢(上)';
if (!workbook.SheetNames.includes(sheetName)) {
  console.error(`Sheet "${sheetName}" 找不到！`);
  console.log('可用的Sheets:', workbook.SheetNames);
  process.exit(1);
}

const worksheet = workbook.Sheets[sheetName];

// 轉換為JSON
const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

// 過濾空行並清理 __EMPTY 列
const jsonData = rawData
  .filter(row => {
    // 過濾掉只有 __EMPTY 的列
    const hasContent = Object.entries(row).some(([key, val]) => {
      return !key.includes('__EMPTY') && val !== '' && val !== null && val !== undefined;
    });
    return hasContent;
  })
  .map(row => {
    // 刪除空列
    const cleaned = {};
    Object.entries(row).forEach(([key, val]) => {
      if (!key.includes('__EMPTY')) {
        cleaned[key] = val;
      }
    });
    return cleaned;
  });

// 輸出到文件
const outputPath = path.join(__dirname, 'output.json');
fs.writeFileSync(outputPath, JSON.stringify(jsonData, null, 2), 'utf-8');

console.log(`✓ 轉換成功！`);
console.log(`Sheet: ${sheetName}`);
console.log(`列數: ${jsonData.length}`);
console.log(`輸出文件: ${outputPath}`);

// 顯示前兩筆記錄作為示例
console.log('\n示例數據 (前2筆):');
console.log(JSON.stringify(jsonData.slice(0, 2), null, 2));
