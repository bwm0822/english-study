const XLSX = require('xlsx');
const wb = XLSX.readFile('./docs/中文.xlsx');

console.log('=== Excel 文件結構 ===\n');
console.log('所有 Sheets:', wb.SheetNames);

wb.SheetNames.forEach(sheetName => {
    const ws = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { defval: '' });

    console.log(`\n【 ${sheetName} 】`);
    console.log(`  總行數: ${data.length}`);

    if (data.length > 0) {
        console.log(`  欄位名: ${Object.keys(data[0]).join(', ')}`);

        // 檢查特殊標記列
        const allKeys = new Set();
        data.forEach(row => {
            Object.keys(row).forEach(k => allKeys.add(k));
        });

        const markers = Array.from(allKeys).filter(k => k.startsWith('#'));
        console.log(`  標記列: ${markers.length > 0 ? markers.join(', ') : '(無)'}`);

        // 列出前3行的成語/標題
        console.log(`  範例:`);
        data.slice(0, 3).forEach((row, i) => {
            const keys = Object.keys(row).filter(k => !k.startsWith('#') && !k.includes('__EMPTY'));
            const values = keys.map(k => `${k}="${row[k].substring(0, 15)}..."`).join(' | ');
            console.log(`    ${i + 1}. ${values}`);
        });
    }
});
