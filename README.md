# xlsx
实现纯前端Excel的导出

目前对单元格样式制定尚未完善 后期考虑做word文档的导出

使用方法见index.html

```js
/**
 * 生成一个新的workbook
 */
let wb = XLSX.utils.book_new();
/**
 * 将二维数组转换成表格
 */
let ws = XLSX.utils.aoa_to_sheet(
  [
    [1, null, '吉米'],
    [4, 5, 6]
  ]
);
/**
 * 设置合并的单元格
 */
ws.merge.push('A1:B1');
/**
 * 将表格与sheet名字插入workbook
 */
XLSX.utils.book_append_sheet(wb, ws, "sheet12");
/**
 * 生成Excel文件
 */
XLSX.writeFile(wb, '测试.xlsx');
```
