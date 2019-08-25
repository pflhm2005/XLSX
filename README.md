# xlsx
实现纯前端Excel的导出

目前支持简单的单元格样式定制 后期考虑做word文档的导出

使用方法如下
```js
// 在测试导出word
let XLSX = Office.XLSX;
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
 * 设置A1,C1单元格的样式
 */
ws.A1.s = {
  fontSize: 14,
  fontWeight: 'bold',
  fontFamily: '微软雅黑',
  textAlign: 'center',
  verticalAlign: 'center',
};
ws.C1.s = {
  fontSize: 12,
  textAlign: 'right',
  verticalAlign: 'bottom',
};
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

#### 样式定制

key|描述|type|可选值|默认值
--|--|--|--|--
fontSize|字体大小|Number|--|12
fontWeight|是否加粗|String|normal,bold|normal
fontFamily|字体类型|String|等线,微软雅黑(尽量不要传其他类型)|等线
textAlign|水平对齐|String|left,right,center|left
verticalAlign|垂直对齐|String|top,bottom,center|top

其他样式基本也不会用到，就不搞了
