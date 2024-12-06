import JSZip from '../dep/jszip';
import GenericOffice from '../GenericOffice';
import JimmyMap from '../JimmyMap';

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
  generateSheetRelsAst,
  generateDrawingRelsPathAst,
  generateDrawingAst
} from './ast';

import {
  columnToNum,
  numToColumn,
  escapeHTML,
  transferCellPos,
} from './util';

/**
 * excel导出类
 */
class XLSX extends GenericOffice {
  constructor() {
    super('xlsx');
    this.sharedStringTotal = 0;
    // 字体映射表默认值
    this.defaultFontAst = generatefontAst('等线', 12, null);
    // 样式映射表默认值
    this.defaultCellXfsAst = generatecellXfsAst(0);
    // Map映射表
    this.sharedStringMap = new JimmyMap(null, false);
    this.fontMap = new JimmyMap(this.defaultFontAst);
    this.cellXfsMap = new JimmyMap(this.defaultCellXfsAst);
    // 单元格的默认样式
    this.defaultStyle = {
      fontFamily: '等线',
      fontSize: 12,
      fontWeight: 'normal',
      textAlign: 'left',
      verticalAlign: 'top',
      wrapText: null,
    };
    // 合法值
    this.verticalDirction = ['top', 'bottom', 'center'];
    this.horizontalDirction = ['left', 'right', 'center'];

    this.defaultConfig = {
      cacheString: true,
    };
    this.imageIndex = 1;
  }
  /**
   * 字体样式重合度较高 直接复用
   * 字符串的重合度也较高 这个方法直接废弃
   */
  cleanUp() {
    this.sharedStringTotal = 0;
    this.sharedStringMap.clearUp();
  }
  /**
   * 直接生成文件
   */
  async writeFile(wb, filename = "未命名.xlsx", config = {}) {
    this.exportFile(await this.write_zip(wb, config), filename);
  }
  /** 
   * 生成xml文件
   */
  async write_zip(wb, config) {
    let zip = new JSZip();
    this.config = Object.assign({}, this.defaultConfig, config);

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
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', Target: workbookXmlPath },
      { Type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', Target: coreXmlPath },
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', Target: appXmlPath },
    ];
    let relsPath = '_rels/.rels';
    zip.file(relsPath, this.writeXml(generateRelsAst(rels)));
    // xl/_rels/workbook.xml.rels
    let rels2 = SheetNames.map((name, i) => {
      return { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', Target: `worksheets/sheet${i + 1}.xml` };
    }).concat([
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', Target: 'theme/theme1.xml' },
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', Target: 'styles.xml' },
      { Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings', Target: 'sharedStrings.xml' },
    ]);
    let xmlRelsPath = 'xl/_rels/workbook.xml.rels';
    zip.file(xmlRelsPath, this.writeXml(generateRelsAst(rels2)));

    // [Content_Types].xml
    let ctXmlPath = '[Content_Types].xml';
    zip.file(ctXmlPath, this.writeXml(generateCtAst(SheetNames)));

    // xl/theme/theme1.xml
    let themeXmlPath = 'xl/theme/theme1.xml';
    zip.file(themeXmlPath, this.writeXml(generateThemeAst()));

    // sheet.xml
    for (let i = 0; i < SheetNames.length; i++) {
      let sheetPath = `xl/worksheets/sheet${i + 1}.xml`;
      let sheetName = SheetNames[i];
      let sheetAst = await this.parseSheet(wb.Sheets[sheetName]);
      const sil = sheetAst.typeObj.imageList.length;
      if (sil) {
        let sheetRelsPath = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`;
        zip.file(sheetRelsPath, this.writeXml(generateSheetRelsAst(i + 1)));
        let drawingRelsPath = `xl/drawings/_rels/drawing${i + 1}.xml.rels`;
        zip.file(drawingRelsPath,  this.writeXml(generateDrawingRelsPathAst(sil, this.imageIndex)));
        let drawingPath = `xl/drawings/drawing${i + 1}.xml`;
        const imageList = sheetAst.typeObj.imageList;
        imageList.forEach((img, i) => {
          let imagePath = `xl/media/image${this.imageIndex + i}.png`;
          zip.file(imagePath, img.arrayBuffer);
        });
        zip.file(drawingPath, this.writeXml(generateDrawingAst(imageList)));
        this.imageIndex += sil;
      }
      zip.file(sheetPath, this.writeXml(generateSheetAst(sheetAst)));
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

    return zip.generate();
  }

  /**
   * 处理SheetAst
   * 由于这里会同时生成style与string的映射map
   * 放到原型方法上处理
   */
  async parseSheet(sheet) {
    // 表格默认从A1开始 所以只计算后面的值
    let ref = sheet.ref || 'A1:A1';
    let range = ref.split(':')[1];
    let len = range.length, r = 0, c = 0;
    // 单元格公式
    let typeObj = {
      list: [],
      imageList: []
    };
    // 计算得到数据的最大行列值
    for (let i = 0; i < len; i++) {
      let unicode = range.charCodeAt(i) - 64;
      if (unicode < 0 || unicode > 26) {
        c = columnToNum(range.slice(0, i));
        r = Number(range.slice(i));
        break;
      }
    }
    // 描述表格数据的数组
    let SheetData = [];

    /**
     * 生成列属性 格式如下
     * <cols>
     *  <col />
     * </cols>
     */
    let cols = [];
    const columnStyleMap = sheet.columnStyles;
    const columnStyleKeys = Object.keys(columnStyleMap);
    const l = columnStyleKeys.length;
    if (l) {
      for (let i = 0; i < l; i++) {
        const key = columnStyleKeys[i];
        const styleConfig = columnStyleMap[key];
        let config = {};
        Object.keys(styleConfig).forEach(attr => {
          if (attr === 'len' || attr === 'height') {
            config.width = styleConfig[attr];
            config.customWidth = '1';
          }
          if (attr === 'hidden') {
            config.hidden = '1';
          }
        });
        cols.push({ n: 'col', p: { min: key, max: key, ...config } });
      }
    }

    /**
     * 生成行 格式如下
     * <row r="1" span="1:?"></row>
     */
    const rowStyleMap = sheet.rowStyles;

    for (let i = 1; i <= r; i++) {
      let rowStyle = rowStyleMap[i];
      let extraParams = {};
      if (rowStyle) {
        extraParams = this.parseRowStyle(rowStyle);
      }

      let rowAst = { n: 'row', p: { r: i, spans: `1:${c}`, ...extraParams }, c: [] };
      let rowChildren = rowAst.c;
      /**
       * 生成列 内容插入row标签中
       */
      for (let j = 0; j < c; j++) {
        let pos = `${numToColumn(j)}${i}`;
        // 默认不生成值为null的单元格 好像也不会出问题
        let cell = sheet[pos];
        if (cell) {
          cell.pos = pos;
          let s = 0;
          if (cell.s) s = this.parseStyle(cell.s);
          // 解析公式格式 特殊公式强制更改单元格值
          if (cell.t) this.parseType(typeObj, cell);
          // 处理公式
          let f = null;
          if (cell.f) {
            f = { n: 'f', t: cell.f };
            cell.v = this.calcFormula(sheet, cell.f);
          }
          let val = cell.v;
          /**
           * warning 
           * 当数据类型与t属性不同步时 文档会报错
           * 然而生成sharedString后 插入值都是数字
           * 因此必须区分类型
           */
          let t = typeof val === 'number' ? 'n' : 's';
          if (t === 's') val = this.LookOrInsertStringMap(cell.v);
          // 处理插入图片
          if (cell.image) {
            await this.parseImage(typeObj, sheet, pos, cell.image);
          }
          rowChildren.push({ n: 'c', p: { r: pos, t, s }, c: [f, { n: 'v', t: val }] });
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
    let mergeAst = null;
    if (merge.length) {
      mergeAst = { n: 'mergeCells', p: { count: merge.length }, c: merge.map(ref => { return { n: 'mergeCell', p: { ref } } }) };
    }

    return { ref, SheetData, mergeAst, cols, typeObj };
  }
  parseRowStyle(rowStyle) {
    let result = {};
    Object.keys(rowStyle).forEach(attr => {
      if (attr === 'len' || attr === 'height') {
        result.ht = rowStyle[attr];
        result.customHeight = '1';
      }
      if (attr === 'hidden') {
        result.hidden = '1';
      }
    });
    return result;
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
    return this.LookOrInsertStyleMap(style);
  }

  /**
   * 解析单元格type
   * list 下拉列表
   * fnc 公式
   */
  parseType(typeObj, cell) {
    const type = cell.t.type;
    if (type === 'list') {
      cell.v = cell.t.list[0];
      typeObj.list.push({
        n: 'dataValidation',
        p: {
          type: 'list',
          allowBlank: '1',
          showInputMessage: '1',
          showErrorMessage: '1',
          sqref: cell.pos,
        },
        c: [{ n: 'formula1', t: `&quot;${cell.t.list.join(',')}&quot;` }]
      });
    }
  }
  /**
   * 解析单元格公式 SUM(A1:D2)
   */
  calcFormula(sheet, formula) {
    const ar = formula.split('(');
    const fuc = ar[0];
    const cellValues = this.getMultiCell(ar[1].slice(0, -1)).map(pos => {
      if (!sheet[pos]) return '';
      return sheet[pos].v;
    }).filter(v => typeof v === 'number');
    const l = cellValues.length;
    const sum = cellValues.reduce((a, b) => a + b);
    switch(fuc) {
      case 'SUM':
        return sum;
      case 'COUNT':
        return l;
      case 'AVERAGE':
        return sum / l;
      case 'MAX':
        return Math.max(cellValues);
      case 'MIN':
        return Math.min(cellValues);
      default:
        throw new Error('未定义的函数类型');
    }
  }
  async parseImage(typeObj, sheet, pos, imageConfig) {
    let imageData = null;
    if (imageConfig.type === 'require') {
      imageData = await this.getImageDataBySrc(imageConfig);
    } else if (imageConfig.type === 'url') {
      imageData = await this.getImageDataByUrl(imageConfig);
    } else if (imageConfig.type === 'upload') {
      imageData = await this.getImageDataByFile(imageConfig);
    }
    
    const { columnStyles, rowStyles } = sheet;
    const posAr = pos.split(/(\d+)/);
    let columnStartIndex = columnToNum(posAr[0]);
    let rowStartIndex = Number(posAr[1]);
    /**
     * 导出的文档默认宽10 高为16
     * 列宽度默认为89px 设置参数时100 = 132px
     * 行高度默认为21px 设置参数时100 = 134px
     * 当手动设置行列长度时 导入图片会被拉伸 这里做手动处理保证图片宽高
     */
    let width = imageData.width;
    let tempLen = 89;
    let i = columnStartIndex;
    while(width > 0) {
      if (columnStyles[i] && columnStyles[i].len) {
        tempLen = Math.round(columnStyles[i].len * 1.32 * 6.1);
      } else {
        tempLen = 89;
      }
      width -= tempLen;
      i++;
    }
    i--;
    // 参数来源实体文件调试 实际公式难搞
    let columnTailOff = (width + tempLen) * 10305;
    if (columnTailOff === 0) i++;
    imageData.xdrCol = [columnStartIndex - 1, i - 1, 635, columnTailOff];
    
    let height = imageData.height;
    tempLen = 21;
    i = rowStartIndex;
    while(height > 0) {
      if (rowStyles[i] && rowStyles[i].len) {
        tempLen = rowStyles[i].len * 1.34;
      } else {
        tempLen = 21;
      }
      height -= tempLen;
      i++
    }
    i--;
    let rowTailOff = (height + tempLen) * 9588;
    imageData.xdrRow = [rowStartIndex - 1, i - 1, 635, rowTailOff];
    typeObj.imageList.push(imageData);
  }
  getImageDataBySrc(imageConfig) {
    return new Promise(resolve => {
      const { src, scale } = imageConfig;
      const img = document.createElement('img');
      img.setAttribute('crossOrigin', 'anonymous');
      img.src = src;
      img.onload = () => {
        this.canvas.width = img.width * scale;
        this.canvas.height = img.height * scale;
        this.ctx.drawImage(img, 0, 0, img.width * scale, img.height * scale);
        this.canvas.toBlob(async (b) => {
          resolve({
            width: img.width,
            height: img.height,
            arrayBuffer: await b.arrayBuffer()
          });
      })};
    });
  }
  getImageDataByFile(imageConfig) {
    return new Promise(resolve => {
      const { file, scale } = imageConfig;
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = (e) => {
        const src = e.target.result;
        const img = document.createElement('img');
        img.src = src;
        img.onload = async () => {
          resolve({
            width: img.width * scale,
            height: img.height * scale,
            arrayBuffer: await file.arrayBuffer()
          });
        };
      }
    });
  }
  getImageDataByUrl(imageConfig) {
    return fetch(imageConfig.url).then(async r => {
      return this.getImageDataByFile({
        scale: imageConfig.scale,
        type: 'upload',
        file: await r.blob()
      });
    }).catch(e => {
      reject(new Error(e));
    });
  }

  /**
   * 不同的对象并不相等
   * 这里直接转换为JSON字符串插入到map中
   */
  LookOrInsertStringMap(val) {
    this.sharedStringTotal++;
    return this.sharedStringMap.LookOrInsert(val);
  }
  LookOrInsertFontMap(style) {
    let { fontFamily, fontSize, fontWeight } = style;
    let fontAst = generatefontAst(fontFamily, fontSize, fontWeight);
    return this.fontMap.LookOrInsert(fontAst);
  }
  LookOrInsertStyleMap(style) {
    style.fontId = this.LookOrInsertFontMap(style);
    let { textAlign, verticalAlign } = style;
    if (!this.horizontalDirction.includes(textAlign)) style.textAlign = 'left';
    if (!this.verticalDirction.includes(verticalAlign)) style.verticalAlign = 'top';
    if (style.wrap) style.wrapText = '1';
    let cellXfsAst = generatecellXfsAst(style);
    return this.cellXfsMap.LookOrInsert(cellXfsAst);
  }

  /**
   * 工具方法
   */
  getMultiCell(str) {
    const ar = str.split(':').join('').split(/(\d+)/).filter(v => v);
    const columnAr = [columnToNum(ar[0]), columnToNum(ar[2])].sort((a, b) => a - b);
    const rowAr = [Number(ar[1]), Number(ar[3])].sort((a, b) => a - b);
    let result = [];
    for(let i = columnAr[0];i <= columnAr[1];i++) {
      const y = numToColumn(i - 1);
      for(let j = rowAr[0];j <= rowAr[1];j++) {
        result.push(y + j);
      }
    }
    return result;
  }
  getColumnRange(start, end, n) {
    let s = start.charCodeAt();
    let e = end.charCodeAt();
    let result = [];
    while (s <= e) {
      result.push(String.fromCharCode(s) + n);
      s++;
    }
    return result;
  }
  book_new() {
    return { SheetNames: [], Sheets: {} };
  }
  aoa_to_sheet(ar) {
    let ws = {
      ref: '',
      merge: [],
      rowStyles: {},
      columnStyles: {},
    };
    let len = ar.length, r = 0, c = 0;
    for (; r < len; r++) {
      for (c = 0; c < ar[r].length; c++) {
        let cell = { v: ar[r][c] };
        if (cell.v === null) continue;
        // else cell.t = 's';
        ws[transferCellPos(r, c)] = cell;
      }
    }
    // 行 => 1,2,3,4,...
    const maxRow = r;
    // 列 => A,B,C,D,...
    const maxColumn = Math.max(...ar.map(v => v.length));
    const tailPos = transferCellPos(maxRow - 1, maxColumn - 1);
    ws.ref = `A1:${tailPos}`;
    ws.arrayInfo = { maxRow, maxColumn, ar };
    return ws;
  }
  setRCStyle(ws, pos, attr, val) {
    let type = 'column';
    let idx = pos;
    if (typeof pos === 'number' || Number(pos) === 'number') {
      type = 'row';
    } else {
      idx = columnToNum(pos);
    }
    let key = type + 'Styles';
    let styleObj = ws[key];
    if (!styleObj[idx]) styleObj[idx] = {};
    if (attr === 'len') {
      if (type === 'column') val = (val / 6.1).toFixed(1);
      if (typeof val === 'number') styleObj[idx][attr] = String(val);
      if (typeof val === 'string') styleObj[idx][attr] = val;
    }
    if (attr === 'hidden') {
      styleObj[idx][attr] = true;
    }
  }
  setImage(ws, pos, img) {
    img.scale = img.scale || 1;
    if (!ws[pos]) {
      const posAr = pos.split(/(\d+)/);
      let posColumn = columnToNum(posAr[0]);
      let posRow = Number(posAr[1]);
      const arrayInfo = ws.arrayInfo;
      let maxColumn = Math.max(arrayInfo.maxColumn, posColumn);
      let maxRow = Math.max(arrayInfo.maxRow, posRow);
      // 进行全面补足
      for(let i = 0; i <= maxColumn; i++) {
        for(let j = 0; j <= maxRow; j++) {
          const curPos = transferCellPos(j - 1, i - 1);
          if (!ws[curPos]) ws[curPos] = { v: '' };
        }
      }
      ws.ref = `A1:${transferCellPos(maxRow - 1, maxColumn - 1)}`;
    }
    ws[pos].image = img;
  }

  book_append_sheet(wb, ws, name = '') {
    if (!name) {
      let len = wb.SheetNames.length;
      name = `Sheet${len + 1}`;
    }
    name = escapeHTML(name);
    wb.SheetNames.push(name);
    wb.Sheets[name] = ws;
  }

  /**
   * 废弃方法
   */
  setRowOrColumnStyle(ws, type, idx, attr, val) {
    if (['row', 'column'].includes(type)) {
      let key = type + 'Styles';
      let styleObj = ws[key];
      if (type === 'column') {
        if (typeof idx === 'string') idx = columnToNum(idx);
      }
      if (!styleObj[idx]) styleObj[idx] = {};
      // 高度
      if (attr === 'height') {
        if (type === 'column') val = (val / 6.1).toFixed(1);
        if (typeof val === 'number') styleObj[idx][attr] = String(val);
        if (typeof val === 'string') styleObj[idx][attr] = val;
      }
      // 隐藏
      if (attr === 'hidden') {
        styleObj[idx][attr] = true;
      }
    }
  }
}

export default new XLSX();