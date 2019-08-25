import JSZip from './dep/jszip';

import {
  generateAppAst,
  generateCoreAst,
  generateRelsAst,
  generateCtAst,
  generateWorkBookAst,
  generatefontAst,
  generatecellXfsAst,
  generateStyleAst,
  generateThemeAst,
  generateSheetAst,
  generateSharedStringAst,
} from './ast';

import {
  columnToNum,
  numToColumn,
  escapeHTML,
  transferCellPos,
} from './util';

const XML_ROOT_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
/**
 * excel导出类
 */
class XLSX {
  constructor() {
    this.a = document.createElement('a');
    /**
     * sharedStringMap映射表
     * 每次导出后需要做重置
     */
    this.sharedStringUid = -1;
    this.sharedStringTotal = 0;
    this.sharedStringMap = new Map();
    // 字体映射表 有一个默认值
    this.defaultFontAst = generatefontAst('等线', 12, null);
    this.fontMap = new Map([[JSON.stringify(this.defaultFontAst), 0]]);
    this.fontUid = 0;
    // 单元格的默认样式
    this.defaultStyle = {
      fontFamily: '等线',
      fontSize: 12,
      fontWeight: 'normal',
      textAlign: 'left',
      verticalAlign: 'top'
    };
    // 合法值
    this.verticalDirction = ['top', 'bottom', 'center'];
    this.horizontalDirction = ['left', 'right', 'center'];
    // 样式映射表 有一个默认值
    this.defaultCellXfsAst = generatecellXfsAst(0);
    this.cellXfsMap = new Map([[JSON.stringify(this.defaultCellXfsAst), 0]]);
    this.cellXfsUid = 0;
  }
  /**
   * 只对字符串的映射表做重置
   * 由于字体样式重合度较高 直接复用
   * TODO 可能字符串的重合度也较高 这个方法直接废弃
   */
  cleanUp() {
    this.sharedStringUid = -1;
    this.sharedStringTotal = 0;
    this.sharedStringMap.clear();
  }
  s2ab(s) {
    let buf = new ArrayBuffer(s.length);
    let view = new Uint8Array(buf);
    for (let i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }
  writeTag(o) {
    if(!o) return '';
    let properties = '';
    if(o.p) properties = Object.keys(o.p).map(key => ` ${key}="${o.p[key]}"`).join('');
    if(o.t === undefined && !o.c) return `<${o.n}${properties}/>`;
    else if(o.c !== undefined) return `<${o.n}${properties}>${(o.c || []).map(v => this.writeTag(v)).join('')}</${o.n}>`;
    else return `<${o.n}${properties}>${o.t}</${o.n}>`;
  }
  writeXml(content) {
    return `${XML_ROOT_HEADER}${this.writeTag(content)}`;
  }
  /**
   * 返回二进制文件流
   */
  write(wb, opt = {}){
    let zip = this.write_zip(wb, opt);
    /**
     * 这里的样式可以考虑复用
     * 字符串清空
     */
    this.cleanUp();
    return this.s2ab(zip.generate({ type: 'string' }));
    // return zip.generateAsync({type:'string'}).then(str => this.s2ab(str));
  }
  /**
   * 直接生成文件
   */
  writeFile(wb, filename = "未命名.xlsx") {
    let blob = new Blob([this.write(wb)], { type: "application/octet-stream" });
    let url = URL.createObjectURL(blob);
    this.a.href = url;
    this.a.download = filename;
    this.a.click();
    setTimeout(() => {
      URL.revokeObjectURL(url);
    }, 10000);
  }
  /** 
   * 生成xml文件
   */
  write_zip(wb) {
    let zip = new JSZip();
    let SheetNames = wb.SheetNames;
    // docProps/app.xml
    let appXmlPath = 'docProps/app.xml';
    zip.file(appXmlPath, this.writeXml(generateAppAst(SheetNames)));

    // docProps/core.xml
    let coreXmlPath = 'docProps/core.xml';
    zip.file(coreXmlPath, this.writeXml(generateCoreAst()));

    // xl/workbook.xml
    let workbookXmlPath = 'xl/workbook.xml';
    zip.file(workbookXmlPath, this.writeXml(generateWorkBookAst(SheetNames)));

    // _rels/.rels
    let rels = [
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', Target: workbookXmlPath  },
      { Type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', Target: coreXmlPath },
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', Target: appXmlPath },
    ];
    let relsPath = '_rels/.rels';
    zip.file(relsPath,this.writeXml(generateRelsAst(rels)));
    // xl/_rels/workbook.xml.rels
    let xmlRels = SheetNames.map((name, i) => {
      return { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', Target: `worksheets/sheet${i+1}.xml` };
    }).concat([
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', Target: 'theme/theme1.xml' },
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', Target: 'styles.xml' },
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings', Target: 'sharedStrings.xml' },
    ]);
    let xmlRelsPath = 'xl/_rels/workbook.xml.rels';
    zip.file(xmlRelsPath, this.writeXml(generateRelsAst(xmlRels)));

    // [Content_Types].xml
    let ctXmlPath = '[Content_Types].xml';
    zip.file(ctXmlPath,this.writeXml(generateCtAst(SheetNames)));

    // xl/theme/theme1.xml
    let themeXmlPath = 'xl/theme/theme1.xml';
    zip.file(themeXmlPath, this.writeXml(generateThemeAst()));

    // sheet.xml
    for(let i = 0;i < SheetNames.length;i++) {
      let sheetPath = `xl/worksheets/sheet${i+1}.xml`;
      let sheetName = SheetNames[i];
      let { ref, SheetData, mergeAst } = this.parseSheet(wb.Sheets[sheetName])
      zip.file(sheetPath, this.writeXml(generateSheetAst(ref, SheetData, mergeAst)));
    }

    /**
     * 字符串与样式的xml依赖于map
     * 必须放到最后
     */
    // xl/styles.xml
    let styleXmlPath = 'xl/styles.xml';
    zip.file(styleXmlPath, this.writeXml(generateStyleAst(this.fontMap, this.cellXfsMap)));

    // xl/sharedStrings.xml
    let sharedStringXmlPath = 'xl/sharedStrings.xml';
    zip.file(sharedStringXmlPath, this.writeXml(generateSharedStringAst(this.sharedStringMap, this.sharedStringTotal)))

    return zip;
  }
  /**
   * 处理SheetAst
   * 由于这里会同时生成style与string的映射map 
   * 放到原型方法上处理
   */
  parseSheet(sheet) {
    // 表格默认从A1开始 所以只计算后面的值
    let ref = sheet.ref || 'A1:A1';
    let range = ref.split(':')[1];
    let len = range.length, r = 0, c = 0;
    // 计算得到数据的最大行列值
    for(let i = 0;i < len;i++) {
      let unicode = range.charCodeAt(i) - 64;
      if(unicode < 0 || unicode > 26) {
        c = columnToNum(range.slice(0, i));
        r = Number(range.slice(i));
        break;
      }
    }
    // 描述表格数据的数组
    let SheetData = [];
    /**
     * 生成行 格式如下
     * <row r="1" span="1:?"></row>
     */
    for(let i = 1;i <= r;i++) {
      let rowAst = { n: 'row', p: { r: i, spans: `1:${c}` }, c: [] };
      let rowChildren = rowAst.c;
      /**
       * 生成列 内容插入row标签中
       */
      for(let j = 0;j < c;j++) {
        let pos = `${numToColumn(j)}${i}`;
        // 默认不生成值为null的单元格 好像也不会出问题
        let cell = sheet[pos];
        if(cell) {
          let s = 0;
          if(cell.s) s = this.parseStyle(cell.s);
          let val = cell.v;
          /**
           * warning 
           * 当数据类型与t属性不同步时 文档会报错
           * 然而生成sharedString后 插入值都是数字
           * 因此必须区分类型
           */
          let t = typeof val === 'number' ? 'n' : 's';
          if(t === 's') val = this.LookOrInsertStringMap(cell.v);
          rowChildren.push({ n: 'c', p: { r: pos, t, s }, c: [{ n: 'v', t: val }] });
        } else {
          rowChildren.push({ n: 'c', p: { r: pos } });
        }
      }
      SheetData.push(rowAst);
    }
    /**
     * 处理单元格合并
     */
    let merge = sheet.merge || [];
    let mergeAst = { n: 'mergeCells', p: { count: merge.length }, c: merge.map(ref => { return { n: 'mergeCell', p: {ref} } }) };

    return { ref, SheetData, mergeAst };
  }
  /**
   * 处理样式映射表
   * 同时返回对应的styleId 默认样式id为0
   * 具体的Ast生成交给generate
   */
  parseStyle(style) {
    /**
     * 需要给默认值
     */
    style = Object.assign({}, this.defaultStyle, style);
    let fontId = this.LookOrInsertFontMap(style);
    return this.LookOrInsertStyleMap(fontId, style);
  }
  /**
   * 不同的对象并不相等
   * 这里直接转换为JSON字符串插入到map中
   */
  LookOrInsert(map, key, val, serialize = true) {
    if(serialize) key = JSON.stringify(key);
    let tar = map.get(key);
    if(tar === undefined) {
      this[val]++;
      map.set(key, this[val]);
      return this[val];
    } else {
      return tar;
    }
  }
  LookOrInsertStringMap(val) {
    this.sharedStringTotal++;
    return this.LookOrInsert(this.sharedStringMap, val, 'sharedStringUid', false);
  }
  LookOrInsertFontMap(style) {
    let { fontFamily, fontSize, fontWeight } = style;
    let fontAst = generatefontAst(fontFamily, fontSize, fontWeight);
    return this.LookOrInsert(this.fontMap, fontAst, 'fontUid');
  }
  LookOrInsertStyleMap(fontId, style) {
    let { textAlign, verticalAlign } = style;
    if(!this.horizontalDirction.includes(textAlign)) textAlign = 'left';
    if(!this.verticalDirction.includes(verticalAlign)) verticalAlign = 'top';
    let cellXfsAst = generatecellXfsAst(fontId, textAlign, verticalAlign);
    return this.LookOrInsert(this.cellXfsMap, cellXfsAst, 'cellXfsUid');
  }

  /**
   * 工具方法
   */
  book_new() {
    return { SheetNames: [], Sheets: {} };
  }
  aoa_to_sheet(ar) {
    let ws = {
      ref: '',
      merge: []
    };
    let len = ar.length, r = 0, c = 0;
    for(;r < len;r++) {
      for(c = 0;c < ar[r].length;c++) {
        let cell = { v: ar[r][c] };
        if(cell.v === null) continue;
        // else cell.t = 's';
        ws[transferCellPos(r, c)] = cell;
      }
    }
    ws.ref = `A1:${transferCellPos(r - 1, Math.max(...ar.map(v => v.length)) - 1)}`;
    return ws;
  }
  book_append_sheet(wb, ws, name = '') {
    if(!name) {
      let len = wb.SheetNames.length;
      name = `Sheet${len+1}`;
    }
    name = escapeHTML(name);
    wb.SheetNames.push(name);
    wb.Sheets[name] = ws;
  }
}

if(typeof exports === 'object' && typeof module !== 'undefined') module.exports = new XLSX();
else window.XLSX = new XLSX();