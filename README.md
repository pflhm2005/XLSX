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
    ['fdsfdfdsafdasfdas', null, '吉米', 'a'],
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
  wrap: true
};
ws.C1.s = {
  fontSize: 12,
  fontWeight: 'bold',
  fontFamily: '微软雅黑',
  textAlign: 'right',
  verticalAlign: 'bottom',
};

// 设置单元格公式
ws.C1.t = {
  type: 'list',
  list: ['吉米', '吉米2', '吉米3'],
};

// 插入图片
XLSX.setImage(ws, 'E5', {
  src: 'http://img.hb.aicdn.com/38d8f519b3f464a80d85ed9632fed904ed0181f41d632-ZHrigO_fw658'
});

// 设置单元格公式
ws.A3.f = 'SUM(A1:A2)';

 // 设置行列高度
XLSX.setRowOrColumnStyle(ws, 'row', 1, 'height', 70);
XLSX.setRowOrColumnStyle(ws, 'column', 'A', 'height', 70);
// 隐藏行列
XLSX.setRowOrColumnStyle(ws, 'column', 'B', 'hidden');

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
wrap|自动换行|Boolean|true/不填|false

<br/>

## 插入图片

注：本功能限制较大 仅支持线上图片且服务器配置了access-control-allow-origin: '*' 办法钻研中

```js
// 示例代码
XLSX.setImage(ws, 'E10', {
  src: 'http://img.hb.aicdn.com/38d8f519b3f464a80d85ed9632fed904ed0181f41d632-ZHrigO_fw658',
  type: 'link',
  scale: 1
});
// file为upload组件上传后的内容 具体使用方法参考index.html中的代码
XLSX.setImage(ws, 'E5', {
  file,
  type: 'upload',
  scale: 0.3
});
```
该代码在表格的指定地点处生成一张图片，左上角为顶点，大小为图片原始尺寸
key|描述|type|可选值|默认值
--|--|--|--|--
type|图片类型|String|link,upload|--
src|当type为link时必填|String|--|--
file|当类型为upload时必填|File|--|--
scale|缩放|Number|0-1|1

## 单元格公式定制

格式值等同于在EXCEL中直接写的字符串，目前支持SUM、COUNT、AVERAGE、MAX、MIN五个公式

注：不支持嵌套公式使用，未做错误处理

### 下拉列表
```
{
  type: 'list',
  list: ['吉米', '吉米2', '吉米3'],  // 下拉选项 默认取第一个作为值
}
```

## 工具方法
<br/>

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
<br/>

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
<br/>

### book_append_sheet(Object wb, Object ws, String name)
> 将工作表的配置添加到Excel基础对象
#### 使用示例
```
// 将内容为ws的工作插入到wb中 sheet的名字是sheet1
book_append_sheet(wb, ws, 'sheet1')
```
#### 无返回
<br/>

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
<br/>

### setRCStyle(Object ws, pos, String attribute, Number|String value)
> 设置行列属性 目前只支持高度、隐藏设置
#### 使用示例
```
// 设置行列高度
XLSX.setRCStyle(ws, 1, 'len', 100);
XLSX.setRCStyle(ws, 'A', 'len', 100);

// 隐藏某一行
XLSX.setRCStyle(ws, 'A', 'hidden');
```
#### 无返回


### <del>setRowOrColumnStyle(Object ws, String type, Number index, String attribute, Number|String value)</del>(已废弃)
> 设置行列属性 目前只支持高度、隐藏设置
#### 使用示例
```
setRowOrColumnStyle(ws, 'row', 1, 'height', 70);
setRowOrColumnStyle(ws, 'column', 'A', 'height', 70);

// 隐藏某一行
setRowOrColumnStyle(ws, 'column', 'A', 'hidden', 70);
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
