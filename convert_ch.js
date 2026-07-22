const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 讀取Excel文件
const filePath = path.join(__dirname, 'docs', '中文.xlsx');
const workbook = XLSX.readFile(filePath);

// 要轉換的sheet列表
const sheetsToConvert = [
  '表現態度1'
];

// 輸出結果結構 - 動態根據sheet名稱創建
const output = {};

let chengYuCount = 0;
let guShiCount = 0;

sheetsToConvert.forEach(sheetName => {
  if (!workbook.SheetNames.includes(sheetName)) {
    console.warn(`⚠ Sheet "${sheetName}" 找不到！`);
    return;
  }

  const worksheet = workbook.Sheets[sheetName];

  // 使用 header:1 模式讀取原始行數據
  const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

  // 用來存儲分類後的數據
  const chengYuData = [];
  const guShiData = [];
  let currentType = null;
  let currentHeaders = [];

  // 遍歷所有行（從第0行開始）
  for (let i = 0; i < rawRows.length; i++) {
    const row = rawRows[i];

    // 檢查是否是類型標記行（#成語 或 #故事）
    if (row[0] === '#成語') {
      currentType = '成語';
      currentHeaders = row.slice(1); // 跳過第一列的標記
      continue;
    } else if (row[0] === '#故事') {
      currentType = '故事';
      currentHeaders = row.slice(1); // 跳過第一列的標記
      continue;
    }

    // 跳過空行
    if (!row.some(cell => cell !== '' && cell !== null && cell !== undefined)) {
      continue;
    }

    // 跳過標記行之前的數據（沒有確定類型）
    if (!currentType) {
      continue;
    }

    // 構建數據對象
    const cleaned = { sheet: sheetName };
    for (let j = 1; j < row.length; j++) {
      const key = currentHeaders[j - 1];
      if (key && row[j] !== '' && row[j] !== null && row[j] !== undefined) {
        cleaned[key] = row[j];
      }
    }

    // 如果有有效內容，添加到相應分類
    const hasValidContent = Object.keys(cleaned).length > 1;
    if (hasValidContent) {
      if (currentType === '成語') {
        chengYuData.push(cleaned);
        chengYuCount++;
      } else if (currentType === '故事') {
        guShiData.push(cleaned);
        guShiCount++;
      }
    }
  }

  // 將數據添加到輸出結構（以sheet名稱為key）
  output[sheetName] = {
    成語: chengYuData,
    故事: guShiData
  };

  console.log(`✓ Sheet: ${sheetName}`);
  console.log(`  - 成語: ${chengYuData.length} 條`);
  console.log(`  - 故事: ${guShiData.length} 條`);
});

// 輸出到文件
const outputPath = path.join(__dirname, 'chinese.json');
fs.writeFileSync(outputPath, JSON.stringify(output, null, 2), 'utf-8');

console.log(`\n✓ 轉換成功！`);
console.log(`\n統計資訊:`);
console.log(`  - 成語: ${chengYuCount} 條`);
console.log(`  - 故事: ${guShiCount} 條`);
console.log(`  - 合計: ${chengYuCount + guShiCount} 條`);
console.log(`\n輸出文件: ${outputPath}`);
console.log(`\n輸出結構:`);
console.log(`{`);
Object.keys(output).forEach(sheetName => {
  console.log(`  "${sheetName}": {`);
  console.log(`    成語: [${output[sheetName].成語.length} 項],`);
  console.log(`    故事: [${output[sheetName].故事.length} 項]`);
  console.log(`  }`);
});
console.log(`}`);
