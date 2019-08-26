import JSZip from '../dep/jszip';
import GenericOffice from '../GenericOffice';

import {
  generateAppAst,
  generateCoreAst,
  generateDocumentAst,
  generateRelsAst,
  generateCtAst,
  generateThemeAst,
  generateWebSettingsAst,
  generateFootTableAst,
  generateSettingsAst,
  generateStylesAst,
} from './ast';

class DOCX extends GenericOffice {
  constructor() {
    super('docx');
    this.rsidRDefault = '001B4009';
    this.defaultFontFamily = 'eastAsia';
    this.defaultStyle = {
      textAlign: 'left',
      fontSize: 24,
      fontWeight: 'normal',
    };
  }
  /**
   * 目前优先支持数组类型
   * @param {Array} ast word文档的抽象语法树
   */
  write(ast) {
    let zip = this.write_zip(ast);
    return zip.generate({ type: 'string' });
  }
  /**
   * 直接生成文件
   */
  writeFile(ast, filename = "未命名.docx") {
    this.exportFile(this.write(ast), filename);
  }
  write_zip(ast) {
    let zip = new JSZip();

    // docProps/app.xml
    let appXmlPath = 'docProps/app.xml';
    zip.file(appXmlPath, this.writeXml(generateAppAst()));

    // docProps/core.xml
    let coreXmlPath = 'docProps/core.xml';
    zip.file(coreXmlPath, this.writeXml(generateCoreAst()));

    let documentXmlPath = 'word/document.xml';
    zip.file(documentXmlPath, this.writeXml(generateDocumentAst(this.parseAst(ast))));

    // _rels/.rels
    let rels = [
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', Target: documentXmlPath  },
      { Type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', Target: coreXmlPath },
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', Target: appXmlPath },
    ];
    let relsPath = '_rels/.rels';
    zip.file(relsPath,this.writeXml(generateRelsAst(rels)));
    let rels2 = [
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', Target: 'styles.xml'  },
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', Target: 'settings.xml' },
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings', Target: 'webSettings.xml' },
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable', Target: 'fontTable.xml' },
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', Target: 'theme/theme1.xml' },
    ];
    let xmlRelsPath = 'word/_rels/document.xml.rels';
    zip.file(xmlRelsPath, this.writeXml(generateRelsAst(rels2)));
    
    // [Content_Types].xml
    let ctXmlPath = '[Content_Types].xml';
    zip.file(ctXmlPath,this.writeXml(generateCtAst()));

    // word/theme/theme1.xml
    let themeXmlPath = 'word/theme/theme1.xml';
    zip.file(themeXmlPath, this.writeXml(generateThemeAst()));

    // word/webSettings.xml
    let webSettingsXmlPath = 'word/webSettings.xml';
    zip.file(webSettingsXmlPath, this.writeXml(generateWebSettingsAst()));

    // word/footTable.xml
    let footTableXmlPath = 'word/footTable.xml';
    zip.file(footTableXmlPath, this.writeXml(generateFootTableAst()));

    // word/settings.xml
    let settingsXmlPath = 'word/settings.xml';
    zip.file(settingsXmlPath, this.writeXml(generateSettingsAst()));

    // word/styles.xml
    let stylesXmlPath = 'word/styles.xml';
    zip.file(stylesXmlPath, this.writeXml(generateStylesAst()));

    return zip;
  }
  /**
   * Ast节点的类型可以是数组、对象、字符串
   * 字符串目前只能是br 代表空行
   * 对象代表正常的文本
   * 数组代表table 参考excel的aoa_to_sheet 会在内部转换为对应的xml
   * @param {Array|Object|String} item 
   */
  parseAst(ast) {
    return ast.map(item => {
      if(typeof item === 'string') return this.parseString(item);
      else if(Array.isArray(item)) return this.parseTable(item);
      else if(typeof item === 'object') return this.parseParagraph(item);
      return null;
    });
  }
  /**
   * 目前对外仅支持br 后续再考虑拓展
   * tab是内部方便我自己用的
   * @param {String} str 特殊标记
   */
  parseString(str, opt = {}) {
    switch(str) {
      case 'br':
        return { n: 'w:p', p: { 'w:rsidRDefault': this.rsidRDefault }, c:[
          { n: 'w:pPr', c: [
            { n: 'w:rPr', c: [
              { n: 'w:sz', p: { 'w:val': 24 } }
            ]}
          ]}
        ]};
      case 'tab':
        return { n: 'w:r', c: [
          { n: 'w:rPr', c: [
            { n: 'w:sz', p: { 'w:val': opt.fz || 24 } }
          ]},
          { n: 'w:tab' }
        ]};
      default:
        throw new Error('unknown Tag');
    }
  }
  /**
   * @param {String} t 文本内容
   * @param {Object} p 文本属性 见下
   * @type {String} textAlign 对齐方式 可选left、center、right
   * @type {Number} fontSize  字体大小
   * @type {String} fontWeight 加粗 bold或不传
   */
  parseParagraph(item) {
    let tAst = { n: 'w:t', t: item.t };
    // 对齐方式
    let pPrAst = null;
    // tab需要自己处理 文本不认这个
    let tabAst = null;
    if(item.t.startsWith('\t')) tabAst = this.parseString('tab');
    let { textAlign, fontSize, fontWeight } = Object.assign({}, this.defaultStyle, item.p);
    // 左对齐与非加粗不做处理 字体类型先不管了
    if(textAlign !== 'left') pPrAst = { n: 'w:pPr', c: [{ n: 'w:jc', p: { 'w:val': textAlign } }] };
    let rPrAst = { n: 'w:rPr', c: [
      { n: 'w:rFonts', p: { 'w:hint': this.defaultFontFamily } },
      fontWeight === 'bold' ? { n: 'w:b' } : null,
      { n: 'w:sz', p: { 'w:val': fontSize } },
      { n: 'w:szCs', p: { 'w:val': fontSize } },
    ]};
    return { n: 'w:p', p: { 'w:rsidRDefault': this.rsidRDefault }, c: [
      pPrAst,
      tabAst,
      { n: 'w:r', c: [rPrAst, tAst]}
    ]};
  }
  /**
   * @param {Array} ar 描述table的二维数组
   * null代表待合并单元格
   * 需要注意的是 由于word表格的复杂性 这个数组必须是一个矩阵
   * 如果不是矩阵 会手动进行修正 并填充为空字符串
   */
  parseTable(ar) {
    let tblPrAst = { n: 'w:tblPr', c: [
      { n: 'w:tblStyle', p: { 'w:val': 'TableGrid' } },
      { n: 'w:tblW', p: { 'w:w': '5000', 'w:type': 'pct' } },
    ]};
    ar = this.checkMatrix(ar);
    return { n: 'w:tbl', c: [tblPrAst, ...ar.map(this.parseRow)] }
  }
  /**
   * 填充不规则数组为矩阵
   */
  checkMatrix(ar) {
    let max = Math.max(...ar.map(v => v.length));
    let flag = ar.some(v => v.length !== max);
    if(!flag) return ar.map(v => {
      if(v.length === max) return v;
      else return v.concat(new Array(max - v.length).fill(''));
    });
    return ar;
  }
  /**
   * 默认加上border
   * @param {Array} row 每一行的数据
   */
  parseRow(row) {
    let tcBorderAst = { n: 'w:tcBorders', c: [
      { n: 'w:top', p: { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': 'auto' } },
      { n: 'w:left', p: { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': 'auto' } },
      { n: 'w:bottom', p: { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': 'auto' } },
      { n: 'w:right', p: { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': 'auto' } },
    ]};
    /**
     * 第一个元素为null逻辑上不合理 设置为空字符串
     * 合并逻辑是向后步进的 这点不做兼容
     * [1, null, null]会被视为3个合并单元格
     * [null, null, 1]则会被视为2个合并单元格与单独的1单元格
     */
    if(row[0] === null) row[0] = '';
    // 检查是否有null值 处理合并单元格
    let needMerge = row.some(v => v === null);
    let GetMergeCellAst = (val) => { return { n: 'w:gridSpan', p: { 'w:val': val } }; };
    let mergeCell = null;
    let len = row.length;
    /**
     * 计算合并单元格
     * 然后进行压缩
     */   
    if(needMerge) {
      let curIndex = 1;
      let nextIndex = 1;
      mergeCell = [];
      while(true) {
        nextIndex = row.slice(curIndex).findIndex(v => v !== null);
        // 代表解析到最后一个元素了
        if(nextIndex === -1) {
          mergeCell.push(len - curIndex);
          break;
        }
        mergeCell.push(nextIndex);
        curIndex = nextIndex + curIndex + 1;
        // 当最后一位非null 直接跳出
        if(curIndex === len) break;
      }
      // 压缩数组 只保留非null数值
      row = row.filter(v => v !== null);
    }
    return { n: 'w:tr', c: row.map((t, i) => {
      return { n: 'w:tc', c: [
        { n: 'w:tcPr', c: [mergeCell ? GetMergeCellAst(mergeCell[i] + 1) : null ,tcBorderAst]},
        { n: 'w:p', c: [
          { n: 'w:r', c: [
            { n: 'w:t', t }
          ]}
        ]}
      ]};
    })};
  }
}

export default new DOCX();