# xlsx
实现纯前端Excel、Word的导出，支持简单的单元格样式定制

使用方法如下

#### 导出Excel
```js
let XLSX = Office.XLSX;
/**
 * 生成一个新的workbook
 */
let wb = XLSX.book_new();
/**
 * 将二维数组转换成表格
 */
let ws = XLSX.aoa_to_sheet(
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
XLSX.book_append_sheet(wb, ws, "sheet12");
/**
 * 生成Excel文件
 */
// XLSX.writeFile(wb, '测试.xlsx');
```

#### 导出Word
```js
let DOCX = Office.DOCX;
// 表格
let table = [
  ['吉米', null, null],
  [1,'是',3],
  [4,'傻逼',6],
];
/**
 * 设置第一行的样式
 */ 
table[0].s = {
  fontWeight: 'bold',
  textAlign: 'center',
};
/**
 * 生成描述docx的抽象语法树
 */
let ast = [
  { t: '测 试', p: { textAlign: 'center', fontSize: 32, fontWeight: 'bold' } },
  'br',
  { t: '这是开头：', p: { fontWeight: 'bold' } },
  { t: '\t这是带了一个tab的段落'},
  table,
];
// DOCX.writeFile(ast, 'word.docx');
```
---

### 导出Excel文档

#### 样式定制

key|描述|type|可选值|默认值
--|--|--|--|--
fontSize|字体大小|Number|--|12
fontWeight|是否加粗|String|normal,bold|normal
fontFamily|字体类型|String|等线,微软雅黑(还支持一些其他值)|等线
textAlign|水平对齐|String|left,right,center|left
verticalAlign|垂直对齐|String|top,bottom,center|top

其他样式基本也不会用到，就不搞了

---
### 导出word文档

目前仅支持AST数组的参数，后续扩展对DOM转换的支持

#### AST

type|描述|可选值
--|--|--
String|特定格式的描述，目前仅支持空行|br
Object|段落的抽象语法树|见下面
Array|Table的二维数组|见下面

##### Object

> 表示一个普通文本

key|type|描述
--|--|--
t|String|段落的文本，若需要tab缩进，直接在前面加\t
p|Object|段落的样式，可选值为textAlign、fontSize、fontWeight，解释见excel导出

##### Array

> 二维数组，表示一个表格，类似于导出excel的aoa_to_sheet
> 
> 与excel不同的是，无需指定mergeCell，值后面出现的null会被强制合并，若需要空表格，请传入空字符串
> 
> 传入的二维数组必须是一个矩阵，不然会被强制填充

> TODO 目前对单元格的样式仅支持水平对齐与粗体 后续不想更新样式了
