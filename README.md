# xlsx
实现纯前端Excel、Word的导出，支持简单的单元格样式定制

使用方法如下

## 导出Excel示例
```js
let XLSX = Office.XLSX;
let wb = XLSX.book_new();
// 生成工作表基本数据
let ws = XLSX.aoa_to_sheet(
  [
    [1, null, '吉米', 'a'],
    [4, 5, 6, 'c']
  ]
);

// 设置合并单元格
ws.merge.push('A1:B1');

// 设置单元格样式
ws.A1.s = {
  fontSize: 14,
  fontWeight: 'bold',
  fontFamily: '微软雅黑',
  textAlign: 'center',
  verticalAlign: 'center',
};
ws.C1.s = {
  fontSize: 12,
  fontWeight: 'bold',
  fontFamily: '微软雅黑',
  textAlign: 'right',
  verticalAlign: 'bottom',
};

// 设置行列的高度
XLSX.setRowOrColumnStyle(ws, 'row', 1, 'height', 70);
XLSX.setRowOrColumnStyle(ws, 'column', 1, 'height', 70);

// 添加工作表到文档对象中
XLSX.book_append_sheet(wb, ws, "测试sheet");
XLSX.writeFile(wb, 'excel.xlsx');
```

## 样式定制

key|描述|type|可选值|默认值
--|--|--|--|--
fontSize|字体大小|Number|--|12
fontWeight|是否加粗|String|normal,bold|normal
fontFamily|字体类型|String|等线,微软雅黑(还支持一些其他值)|等线
textAlign|水平对齐|String|left,right,center|left
verticalAlign|垂直对齐|String|top,bottom,center|top

其他样式基本也不会用到，就不搞了

## 工具方法

### book_new(void void)
> 返回一个xlsx的基础文档对象，具体内容如下
#### 使用示例
```
const wb = book_new()
```
#### 返回值
```
{
  SheetName: [], // 工作表的名字
  Sheets: {} // 工作表映射对象
}
```

### aoa_to_sheet(Array table)
> 将二维数组转换为工作表配置
#### 使用示例
```
const ws = aoa_to_sheet([
  [1,2,3],
  [4,5,6],
  [7,8,9]
])
```
#### 返回值
```
{
  ref: 'A1:C3', // 表格的范围值
  merge: [], // 表格的合并信息
  rowStyle: {}, // 行格式集合
  columnStyle: {},  // 列格式集合
  A1: { v: 1 }  // 每一个表格的值
  A2: { v: 2 },
  ...,
  C3: { v: 9 },
}
```

### book_append_sheet(Object wb, Object ws, String name)
> 将工作表的配置添加到Excel基础对象
#### 使用示例
```
// 将内容为ws的工作插入到wb中 sheet的名字是sheet1
book_append_sheet(wb, ws, 'sheet1')
```
#### 无返回

### getColumnRange(String alpha, String alpha, Number index)
> 返回指定列的数个单元格
#### 使用示例
```
const pos = getColumnRange('A', 'C', 1)
```
#### 返回值
```
['A1', 'B1', 'C1']
```

### setRowOrColumnStyle(Object ws, String type, Number index, String attribute, Number|String value)
> 设置行列属性 目前只支持高度设置
#### 使用示例
```
// 在内部处理时 行高单位为磅而列宽单位为字符 为了方便统一为磅 
// 1字符约等于6.1磅
setRowOrColumnStyle(ws, 'row', 1, 'height', 70);
setRowOrColumnStyle(ws, 'column', 'A', 'height', 70);
```
#### 无返回


---


## 导出Word示例
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
/**
 * 导出word文档
 */
DOCX.writeFile(ast, 'word.docx');
```

## 文档

目前仅支持AST数组的参数，后续扩展对DOM转换的支持

### AST

type|描述|可选值
--|--|--
String|特定格式的描述，目前仅支持空行|br
Object|段落的抽象语法树|见下面
Array|Table的二维数组|见下面

### Object

> 表示一个普通文本

key|type|描述
--|--|--
t|String|段落的文本，若需要tab缩进，直接在前面加\t
p|Object|段落的样式，可选值为textAlign、fontSize、fontWeight，解释见excel导出

#### Array

> 二维数组，表示一个表格，类似于导出excel的aoa_to_sheet
> 
> 与excel不同的是，无需指定mergeCell，值后面出现的null会被强制合并，若需要空表格，请传入空字符串
> 
> 传入的二维数组必须是一个矩阵，不然会被强制填充

> TODO 目前对单元格的样式仅支持水平对齐与粗体 后续不想更新样式了
