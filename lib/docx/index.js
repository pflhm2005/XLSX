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
  }
  write() {
    let zip = this.write_zip(null);
    return zip.generate({ type: 'string' });
  }
  /**
   * 直接生成文件
   */
  writeFile(wb, filename = "未命名.docx") {
    this.exportFile(this.write(), filename);
  }
  write_zip() {
    let zip = new JSZip();

    // docProps/app.xml
    let appXmlPath = 'docProps/app.xml';
    zip.file(appXmlPath, this.writeXml(generateAppAst()));

    // docProps/core.xml
    let coreXmlPath = 'docProps/core.xml';
    zip.file(coreXmlPath, this.writeXml(generateCoreAst()));

    let documentXmlPath = 'word/document.xml';
    zip.file(documentXmlPath, this.writeXml(generateDocumentAst()));

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
}

window.DOCX = new DOCX();