export const generateAppAst = () => {
  return {
    n:'Properties',
    p: {
      xmlns: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
      'xmlns:vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
    },
    c: [
      { n: 'Template', t: 'Normal.dotm' },
      { n: 'TotalTime', t: '0' },
      { n: 'Pages', t: '1' },
      { n: 'Words', t: '0' },
      { n: 'Characters', t: '0' },
      { n: 'Application', t: 'Microsoft Office Word' },
      { n: 'DocSecurity', t: '0' },
      { n: 'Lines', t: '0' },
      { n: 'ScaleCrop', t: 'false' },
      { n: 'Company' },
      { n: 'LinksUpToDate', t: 'false' },
      { n: 'CharactersWithSpaces', t: '0' },
      { n: 'SharedDoc', t: 'false' },
      { n: 'HyperlinksChanged', t: 'false' },
      { n: 'AppVersion', t: '16.0300' },
    ]
  }
};

export const generateCoreAst = () => {
  let date = new Date().toISOString().replace(/\.\d*/, "");
  return {
    n: 'cp:coreProperties',
    p: {
      'xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
      'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
      'xmlns:dcterms': 'http://purl.org/dc/terms/',
      'xmlns:dcmitype': 'http://purl.org/dc/dcmitype/',
      'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    },
    c: [
      { n: 'dc:title' },
      { n: 'dc:subject' },
      { n: 'dc:creator', t: 'feilong pang' },
      { n: 'cp:keywords' },
      { n: 'dc:description' },
      { n: 'cp:lastModifiedBy', t: 'feilong pang' },
      { n: 'cp:revision', t: '1' },
      { n: 'dcterms:created', p: { 'xsi:type': 'dcterms:W3CDTF' }, t: date },
      { n: 'dcterms:modified', p: { 'xsi:type': 'dcterms:W3CDTF' }, t: date },
    ],
  };
};

export const generateDocumentAst = () => {
  return {
    n: 'w:document', p: {
      'xmlns:wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
      'xmlns:cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
      'xmlns:cx1': 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex',
      'xmlns:cx2': 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex',
      'xmlns:cx3': 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex',
      'xmlns:cx4': 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex',
      'xmlns:cx5': 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex',
      'xmlns:cx6': 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex',
      'xmlns:cx7': 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex',
      'xmlns:cx8': 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex',
      'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
      'xmlns:aink': 'http://schemas.microsoft.com/office/drawing/2016/ink',
      'xmlns:am3d': 'http://schemas.microsoft.com/office/drawing/2017/model3d',
      'xmlns:o': 'urn:schemas-microsoft-com:office:office',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'xmlns:m': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
      'xmlns:v': 'urn:schemas-microsoft-com:vml',
      'xmlns:wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
      'xmlns:wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
      'xmlns:w10': 'urn:schemas-microsoft-com:office:word',
      'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'xmlns:w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
      'xmlns:w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
      'xmlns:w16cid': 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
      'xmlns:w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
      'xmlns:wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
      'xmlns:wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
      'xmlns:wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
      'xmlns:wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
      'mc:Ignorable': 'w14 w15 w16se w16cid wp14',
    }, c: [
      { n: 'w:body', c: [
        { n: 'w:p', p: { 'w:rsidR': '004C5786', 'w:rsidRDefault': '00ED5EAE', c: [
          { n: 'w:bookmarkStart', p: { 'w:id': '0', 'w:name': '_GoBack' } },
          { n: 'w:bookmarkEnd', p: { 'w:id': '0' } },
        ]}},
        { n: 'w:sectPr', p: { 'w:rsidR': '004C5786', 'w:rsidSect': '00251A90' }, c: [
          { n: 'w:pgSz', p: { 'w:w': '11900', 'w:h': '16840' } },
          { n: 'w:pgMar', p: { 'w:top': '1440', 'w:right': '1800', 'w:bottom': '1440', 'w:left': '1800', 'w:header': '851', 'w:footer': '992', 'w:gutter': '0' } },
          { n: 'w:cols', p: { 'w:space': '425' } },
          { n: 'w:docGrid', p: { 'w:type': 'lines', 'w:linePitch': '312' } },
        ]}
    ]}]
  };
}

export const generateRelsAst = rels => {
  return {
    n: 'Relationships',
    p: { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' },
    c: rels.map((rel, i) => { return { n: 'Relationship', p: { Id: `rId${i + 1}`, Type: rel.Type, Target: rel.Target } }; })
  };
};

export const generateCtAst = (SheetNames) => {
  return {
    n: 'Types',
    p: { xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types' },
    c: [
      { n: 'Default', p: { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' } },
      { n: 'Default', p: { Extension: 'xml', ContentType: 'application/xml' } },
      { n: 'Override', p: { PartName: '/word/document.xml', ContentType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml' } },
      { n: 'Override', p: { PartName: '/word/styles.xml', ContentType:'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml' } },
      { n: 'Override', p: { PartName: '/word/settings.xml', ContentType:'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml' } },
      { n: 'Override', p: { PartName: '/word/webSettings.xml', ContentType:'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml' } },
      { n: 'Override', p: { PartName: '/word/fontTable.xml', ContentType:'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml' } },
      { n: 'Override', p: { PartName: '/word/theme/theme1.xml', ContentType:'application/vnd.openxmlformats-officedocument.theme+xml' } },
      { n: 'Override', p: { PartName: '/docProps/core.xml', ContentType:'application/vnd.openxmlformats-package.core-properties+xml' } },
      { n: 'Override', p: { PartName: '/docProps/app.xml', ContentType:'application/vnd.openxmlformats-officedocument.extended-properties+xml' } },
    ]
  }
};

export const generateThemeAst = () => {
  return {
    n: 'a:theme',
    p: {
      'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      name: 'Office 主题',
    },
    c: [
      { n: 'a:themeElements', c: [
        { n: 'a:clrScheme', p: { name: 'Office' }, c: [
          { n: 'a:dk1', c: [{ n: 'a:sysClr', p: { val: 'windowText', lastClr: '000000' } }] },
          { n: 'a:lt1', c: [{ n: 'a:sysClr', p: { val: 'window', lastClr: 'FFFFFF' } }] },
          { n: 'a:dk2', c: [{ n: 'a:srgbClr', p: { val: '44546A' } }] },
          { n: 'a:lt2', c: [{ n: 'a:srgbClr', p: { val: 'E7E6E6' } }] },
          { n: 'a:accent1', c: [{ n: 'a:srgbClr', p: { val: '4472C4' } }] },
          { n: 'a:accent2', c: [{ n: 'a:srgbClr', p: { val: 'ED7D31' } }] },
          { n: 'a:accent3', c: [{ n: 'a:srgbClr', p: { val: 'A5A5A5' } }] },
          { n: 'a:accent4', c: [{ n: 'a:srgbClr', p: { val: 'FFC000' } }] },
          { n: 'a:accent5', c: [{ n: 'a:srgbClr', p: { val: '5B9BD5' } }] },
          { n: 'a:accent6', c: [{ n: 'a:srgbClr', p: { val: '70AD47' } }] },
          { n: 'a:hlink', c: [{ n: 'a:srgbClr', p: { val: '0563C1' } }] },
          { n: 'a:folHlink', c: [{ n: 'a:srgbClr', p: { val: '954F72' } }] },
        ]},
        { n: 'a:fontScheme', p: { name: 'Office' }, c:[
          { n: 'a:majorFont', c: [
            { n: 'a:latin', p: { typeface: '等线 Light', panose: '020F0302020204030204' } },
            { n: 'a:ea', p: { typeface: '' } },
            { n: 'a:cs', p: { typeface: '' } },
            { n: 'a:font', p: { script: 'Jpan', typeface: '游ゴシック Light' } },
            { n: 'a:font', p: { script: 'Hang', typeface: '맑은 고딕' } },
            { n: 'a:font', p: { script: 'Hans', typeface: '等线 Light' } },
            { n: 'a:font', p: { script: 'Hant', typeface: '新細明體' } },
            { n: 'a:font', p: { script: 'Arab', typeface: 'Times New Roman' } },
            { n: 'a:font', p: { script: 'Hebr', typeface: 'Times New Roman' } },
            { n: 'a:font', p: { script: 'Thai', typeface: 'Angsana New' } },
            { n: 'a:font', p: { script: 'Ethi', typeface: 'Nyala' } },
            { n: 'a:font', p: { script: 'Beng', typeface: 'Vrinda' } },
            { n: 'a:font', p: { script: 'Gujr', typeface: 'Shruti' } },
            { n: 'a:font', p: { script: 'Khmr', typeface: 'MoolBoran' } },
            { n: 'a:font', p: { script: 'Knda', typeface: 'Tunga' } },
            { n: 'a:font', p: { script: 'Guru', typeface: 'Raavi' } },
            { n: 'a:font', p: { script: 'Cans', typeface: 'Euphemia' } },
            { n: 'a:font', p: { script: 'Cher', typeface: 'Plantagenet Cherokee' } },
            { n: 'a:font', p: { script: 'Yiii', typeface: 'Microsoft Yi Baiti' } },
            { n: 'a:font', p: { script: 'Tibt', typeface: 'Microsoft Himalaya' } },
            { n: 'a:font', p: { script: 'Thaa', typeface: 'MV Boli' } },
            { n: 'a:font', p: { script: 'Deva', typeface: 'Mangal' } },
            { n: 'a:font', p: { script: 'Telu', typeface: 'Gautami' } },
            { n: 'a:font', p: { script: 'Taml', typeface: 'Latha' } },
            { n: 'a:font', p: { script: 'Syrc', typeface: 'Estrangelo Edessa' } },
            { n: 'a:font', p: { script: 'Orya', typeface: 'Kalinga' } },
            { n: 'a:font', p: { script: 'Mlym', typeface: 'Kartika' } },
            { n: 'a:font', p: { script: 'Laoo', typeface: 'DokChampa' } },
            { n: 'a:font', p: { script: 'Sinh', typeface: 'Iskoola Pota' } },
            { n: 'a:font', p: { script: 'Mong', typeface: 'Mongolian Baiti' } },
            { n: 'a:font', p: { script: 'Viet', typeface: 'Times New Roman' } },
            { n: 'a:font', p: { script: 'Uigh', typeface: 'Microsoft Uighur' } },
            { n: 'a:font', p: { script: 'Geor', typeface: 'Sylfaen' } },
            { n: 'a:font', p: { script: 'Armn', typeface: 'Arial' } },
            { n: 'a:font', p: { script: 'Bugi', typeface: 'Leelawadee UI' } },
            { n: 'a:font', p: { script: 'Bopo', typeface: 'Microsoft JhengHei' } },
            { n: 'a:font', p: { script: 'Java', typeface: 'Javanese Text' } },
            { n: 'a:font', p: { script: 'Lisu', typeface: 'Segoe UI' } },
            { n: 'a:font', p: { script: 'Mymr', typeface: 'Myanmar Text' } },
            { n: 'a:font', p: { script: 'Nkoo', typeface: 'Ebrima' } },
            { n: 'a:font', p: { script: 'Olck', typeface: 'Nirmala UI' } },
            { n: 'a:font', p: { script: 'Osma', typeface: 'Ebrima' } },
            { n: 'a:font', p: { script: 'Phag', typeface: 'Phagspa' } },
            { n: 'a:font', p: { script: 'Syrn', typeface: 'Estrangelo Edessa' } },
            { n: 'a:font', p: { script: 'Syrj', typeface: 'Estrangelo Edessa' } },
            { n: 'a:font', p: { script: 'Syre', typeface: 'Estrangelo Edessa' } },
            { n: 'a:font', p: { script: 'Sora', typeface: 'Nirmala UI' } },
            { n: 'a:font', p: { script: 'Tale', typeface: 'Microsoft Tai Le' } },
            { n: 'a:font', p: { script: 'Talu', typeface: 'Microsoft New Tai Lue' } },
            { n: 'a:font', p: { script: 'Tfng', typeface: 'Ebrima' } },
          ]},
          { n: 'a:minorFont', c: [
            { n: 'a:latin', p: { typeface: '等线', panose: '020F0502020204030204' } },
            { n: 'a:ea', p: { typeface: '' } },
            { n: 'a:cs', p: { typeface: '' } },
            { n: 'a:font', p: { script:'Jpan', typeface:"游明朝" } },
            { n: 'a:font', p: { script:'Hang', typeface:"맑은 고딕" } },
            { n: 'a:font', p: { script:'Hans', typeface:"等线" } },
            { n: 'a:font', p: { script:'Hant', typeface:"新細明體" } },
            { n: 'a:font', p: { script:'Arab', typeface:"Arial" } },
            { n: 'a:font', p: { script:'Hebr', typeface:"Arial" } },
            { n: 'a:font', p: { script:'Thai', typeface:"Cordia New" } },
            { n: 'a:font', p: { script:'Ethi', typeface:"Nyala" } },
            { n: 'a:font', p: { script:'Beng', typeface:"Vrinda" } },
            { n: 'a:font', p: { script:'Gujr', typeface:"Shruti" } },
            { n: 'a:font', p: { script:'Khmr', typeface:"DaunPenh" } },
            { n: 'a:font', p: { script:'Knda', typeface:"Tunga" } },
            { n: 'a:font', p: { script:'Guru', typeface:"Raavi" } },
            { n: 'a:font', p: { script:'Cans', typeface:"Euphemia" } },
            { n: 'a:font', p: { script:'Cher', typeface:"Plantagenet Cherokee" } },
            { n: 'a:font', p: { script:'Yiii', typeface:"Microsoft Yi Baiti" } },
            { n: 'a:font', p: { script:'Tibt', typeface:"Microsoft Himalaya" } },
            { n: 'a:font', p: { script:'Thaa', typeface:"MV Boli" } },
            { n: 'a:font', p: { script:'Deva', typeface:"Mangal" } },
            { n: 'a:font', p: { script:'Telu', typeface:"Gautami" } },
            { n: 'a:font', p: { script:'Taml', typeface:"Latha" } },
            { n: 'a:font', p: { script:'Syrc', typeface:"Estrangelo Edessa" } },
            { n: 'a:font', p: { script:'Orya', typeface:"Kalinga" } },
            { n: 'a:font', p: { script:'Mlym', typeface:"Kartika" } },
            { n: 'a:font', p: { script:'Laoo', typeface:"DokChampa" } },
            { n: 'a:font', p: { script:'Sinh', typeface:"Iskoola Pota" } },
            { n: 'a:font', p: { script:'Mong', typeface:"Mongolian Baiti" } },
            { n: 'a:font', p: { script:'Viet', typeface:"Arial" } },
            { n: 'a:font', p: { script:'Uigh', typeface:"Microsoft Uighur" } },
            { n: 'a:font', p: { script:'Geor', typeface:"Sylfaen" } },
            { n: 'a:font', p: { script:'Armn', typeface:"Arial" } },
            { n: 'a:font', p: { script:'Bugi', typeface:"Leelawadee UI" } },
            { n: 'a:font', p: { script:'Bopo', typeface:"Microsoft JhengHei" } },
            { n: 'a:font', p: { script:'Java', typeface:"Javanese Text" } },
            { n: 'a:font', p: { script:'Lisu', typeface:"Segoe UI" } },
            { n: 'a:font', p: { script:'Mymr', typeface:"Myanmar Text" } },
            { n: 'a:font', p: { script:'Nkoo', typeface:"Ebrima" } },
            { n: 'a:font', p: { script:'Olck', typeface:"Nirmala UI" } },
            { n: 'a:font', p: { script:'Osma', typeface:"Ebrima" } },
            { n: 'a:font', p: { script:'Phag', typeface:"Phagspa" } },
            { n: 'a:font', p: { script:'Syrn', typeface:"Estrangelo Edessa" } },
            { n: 'a:font', p: { script:'Syrj', typeface:"Estrangelo Edessa" } },
            { n: 'a:font', p: { script:'Syre', typeface:"Estrangelo Edessa" } },
            { n: 'a:font', p: { script:'Sora', typeface:"Nirmala UI" } },
            { n: 'a:font', p: { script:'Tale', typeface:"Microsoft Tai Le" } },
            { n: 'a:font', p: { script:'Talu', typeface:"Microsoft New Tai Lue" } },
            { n: 'a:font', p: { script:'Tfng', typeface:"Ebrima" } },
          ]}
        ]},
        { n: 'a:fmtScheme', p: { name: 'Office' }, c: [
          { n: 'a:fillStyleLst', c:[
            { n: 'a:solidFill', c: [
              { n: 'a:schemeClr', p: { val: 'phClr' } }
            ]},
            { n: 'a:gradFill', p: { rotWithShape: '1' }, c: [
              { n: 'a:gsLst', c: [
                { n: 'a:gs', p: { pos: '0' }, c: [
                  { n: 'a:schemeClr', p: { val: 'phClr' }, c: [
                    { n: 'a:lumMod', p: { val: '110000' } },
                    { n: 'a:satMod', p: { val: '105000' } },
                    { n: 'a:tint', p: { val: '67000' } },
                  ]}
                ]},
                { n: 'a:gs', p: { pos: '50000' }, c: [
                  { n: 'a:schemeClr', p: { val: 'phClr' }, c: [
                    { n: 'a:lumMod', p: { val: '105000' } },
                    { n: 'a:satMod', p: { val: '103000' } },
                    { n: 'a:tint', p: { val: '73000' } },
                  ]}
                ]},
                { n: 'a:gs', p: { pos: '100000' }, c: [
                  { n: 'a:schemeClr', p: { val: 'phClr' }, c: [
                    { n: 'a:lumMod', p: { val: '105000' } },
                    { n: 'a:satMod', p: { val: '109000' } },
                    { n: 'a:tint', p: { val: '81000' } },
                  ]}
                ]},
              ]},
              { n: 'a:lin', p: { ang: '5400000', scaled: '0' } }
            ]},
            { n: 'a:gradFill', p: { rotWithShape: '1' }, c: [
              { n: 'a:gsLst', c: [
                { n: 'a:gs', p: { pos: '0' }, c: [
                  { n: 'a:schemeClr', p: { val: 'phClr' }, c: [
                    { n: 'a:satMod', p: { val: '103000' } },
                    { n: 'a:lumMod', p: { val: '102000' } },
                    { n: 'a:tint', p: { val: '94000' } },
                  ]}
                ]},
                { n: 'a:gs', p: { pos: '50000' }, c: [
                  { n: 'a:schemeClr', p: { val: 'phClr' }, c: [
                    { n: 'a:satMod', p: { val: '110000' } },
                    { n: 'a:lumMod', p: { val: '100000' } },
                    { n: 'a:shade', p: { val: '100000' } },
                  ]}
                ]},
                { n: 'a:gs', p: { pos: '100000' }, c: [
                  { n: 'a:schemeClr', p: { val: 'phClr' }, c: [
                    { n: 'a:lumMod', p: { val: '99000' } },
                    { n: 'a:satMod', p: { val: '120000' } },
                    { n: 'a:shade', p: { val: '78000' } },
                  ]}
                ]},
              ]},
              { n: 'a:lin', p: { ang: '5400000', scaled: '0' } }
            ]},
          ]},
          { n: 'a:lnStyleLst', c:[
            { n: 'a:ln', p: { w: '6350', cap: 'flat', cmpd: 'sng', algn: 'ctr' }, c:[
              { n: 'a:solidFill', c: [
                { n: 'a:schemeClr', p: { val: 'phClr' } }
              ]},
              { n: 'a:prstDash', p: { val: 'solid' } },
              { n: 'a:miter', p: { lim: '800000' } },
            ]},
            { n: 'a:ln', p: { w: '12700', cap: 'flat', cmpd: 'sng', algn: 'ctr' }, c:[
              { n: 'a:solidFill', c: [
                { n: 'a:schemeClr', p: { val: 'phClr' } }
              ]},
              { n: 'a:prstDash', p: { val: 'solid' } },
              { n: 'a:miter', p: { lim: '800000' } },
            ]},
            { n: 'a:ln', p: { w: '19050', cap: 'flat', cmpd: 'sng', algn: 'ctr' }, c:[
              { n: 'a:solidFill', c: [
                { n: 'a:schemeClr', p: { val: 'phClr' } }
              ]},
              { n: 'a:prstDash', p: { val: 'solid' } },
              { n: 'a:miter', p: { lim: '800000' } },
            ]},
          ]},
          { n: 'a:effectStyleLst', c: [
            { n: 'a:effectStyle', c: [{ n: 'a:effectLst' }]},
            { n: 'a:effectStyle', c: [{ n: 'a:effectLst' }]},
            { n: 'a:effectStyle', c: [
              { n: 'a:effectLst', c: [
                { n: 'a:outerShdw', p: { blurRad: '57150', dist: '19050', dir: '5400000', algn: 'ctr', rotWithShape: '0' }, c:[
                  { n: 'a:srgbClr', p: { val: '000000' },c: [
                    { n: 'a:alpha', p: { val: '63000' } }
                  ]}
                ]},
              ]}
            ]},
          ]},
          { n: 'a:bgFillStyleLst', c: [
            { n: 'a:solidFill', c: [
              { n: 'a:schemeClr', p: { val: 'phClr' } },
            ]},
            { n: 'a:solidFill', c: [
              { n: 'a:schemeClr', p: { val: 'phClr' }, c:[
                { n: 'a:tint', p: { val: '95000' }},
                { n: 'a:satMod', p: { val: '170000' }},
              ]},
            ]},
            { n: 'a:gradFill', p: { rotWithShape: '1' }, c: [
              { n: 'a:gsLst', c: [
                { n: 'a:gs', p: { pos: '0' }, c:[
                  { n: 'a:schemeClr', p: { val: 'phClr' }, c: [
                    { n: 'a:tint', p: { val: '93000' } },
                    { n: 'a:satMod', p: { val: '150000' } },
                    { n: 'a:shade', p: { val: '98000' } },
                    { n: 'a:lumMod', p: { val: '102000' } },
                  ]}
                ]},
                { n: 'a:gs', p: { pos: '50000' }, c:[
                  { n: 'a:schemeClr', p: { val: 'phClr' }, c: [
                    { n: 'a:tint', p: { val: '98000' } },
                    { n: 'a:satMod', p: { val: '130000' } },
                    { n: 'a:shade', p: { val: '90000' } },
                    { n: 'a:lumMod', p: { val: '103000' } },
                  ]}
                ]},
                { n: 'a:gs', p: { pos: '100000' }, c:[
                  { n: 'a:schemeClr', p: { val: 'phClr' }, c: [
                    { n: 'a:shade', p: { val: '63000' } },
                    { n: 'a:satMod', p: { val: '120000' } },
                  ]}
                ]},
              ]},
              { n: 'a:lin', p: { ang: '5400000', scaled: '0' } }
            ]}
          ]}
        ]}
      ]},
      { n: 'a:objectDefaults' },
      { n: 'a:extraClrSchemeLst' },
      { n: 'a:extLst', c:[
        { n: 'a:ext', p: { uri: '{05A4C25C-085E-4340-85A3-A5531E510DB2}'}, c: [
          { n: 'thm15:themeFamily', p: {
            'xmlns:thm15': 'http://schemas.microsoft.com/office/thememl/2012/main',
            name: 'Office Theme',
            id: '{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}',
            vid: '{4A3C46E8-61CC-4603-A589-7422A47A8E4A}',
          }}
        ]}
      ]},
    ]
  };
};

export const generateWebSettingsAst = () => {
  return {
    n: 'w:webSettings',
    p: {
      'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'xmlns:w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
      'xmlns:w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
      'xmlns:w16cid': 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
      'xmlns:w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
      'mc:Ignorable': 'w14 w15 w16se w16cid',
    }, c: [
      { n: 'w:optimizeForBrowser' },
      { n: 'w:allowPNG' },
    ]
  };
};

export const generateFootTableAst = () => {
  return {
    n: 'w:fonts',
    p: {
      'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'xmlns:w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
      'xmlns:w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
      'xmlns:w16cid': 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
      'xmlns:w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
      'mc:Ignorable': 'w14 w15 w16se w16cid',
    }, c: [
      { n: 'w:font', p: { 'w:name': 'DengXian' }, c: [
        { n: 'w:altName', p: { 'w:val': '等线' } },
        { n: 'w:panose1', p: { 'w:val': '02010600030101010101' } },
        { n: 'w:charset', p: { 'w:val': '86' } },
        { n: 'w:family', p: { 'w:val': 'auto' } },
        { n: 'w:pitch', p: { 'w:val': 'variable' } },
        { n: 'w:sig', p: { 'w:usb0': 'A00002BF', 'w:usb1': '38CF7CFA', 'w:usb2': '00000016', 'w:usb3': '00000000', 'w:csb0': '0004000F', 'w:csb1': '00000000' } },
      ]},
      { n: 'w:font', p: { 'w:name': 'Times New Roman' }, c: [
        { n: 'w:panose1', p: { 'w:val': '02020603050405020304' } },
        { n: 'w:charset', p: { 'w:val': '00' } },
        { n: 'w:family', p: { 'w:val': 'roman' } },
        { n: 'w:pitch', p: { 'w:val': 'variable' } },
        { n: 'w:sig', p: { 'w:usb0': '20002A87', 'w:usb1': '80000000', 'w:usb2': '00000008', 'w:usb3': '00000000', 'w:csb0': '000001FF', 'w:csb1': '00000000' } },
      ]},
      { n: 'w:font', p: { 'w:name': '等线 Light' }, c: [
        { n: 'w:panose1', p: { 'w:val': '02010600030101010101' } },
        { n: 'w:charset', p: { 'w:val': '86' } },
        { n: 'w:family', p: { 'w:val': 'auto' } },
        { n: 'w:pitch', p: { 'w:val': 'variable' } },
        { n: 'w:sig', p: { 'w:usb0': 'A00002BF', 'w:usb1': '38CF7CFA', 'w:usb2': '00000016', 'w:usb3': '00000000', 'w:csb0': '0004000F', 'w:csb1': '00000000' } },
      ]}
    ]
  }
};

export const generateSettingsAst = () => {
  return {
    n: 'w:settings', p: {
      'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
      'xmlns:o': 'urn:schemas-microsoft-com:office:office"',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'xmlns:m': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
      'xmlns:v': 'urn:schemas-microsoft-com:vml',
      'xmlns:w10': 'urn:schemas-microsoft-com:office:word',
      'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'xmlns:w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
      'xmlns:w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
      'xmlns:w16cid': 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
      'xmlns:w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
      'xmlns:sl': 'http://schemas.openxmlformats.org/schemaLibrary/2006/main',
      'mc:Ignorable': 'w14 w15 w16se w16cid',
    }, c: [
      { n: 'w:zoom', p: { 'w:percent': '100' } },
      { n: 'w:bordersDoNotSurroundHeader' },
      { n: 'w:bordersDoNotSurroundFooter' },
      { n: 'w:defaultTabStop', p: { 'w:val': '420' } },
      { n: 'w:drawingGridVerticalSpacing', p: { 'w:val': '156' } },
      { n: 'w:displayHorizontalDrawingGridEvery', p: { 'w:val': '0' } },
      { n: 'w:displayVerticalDrawingGridEvery', p: { 'w:val': '2' } },
      { n: 'w:characterSpacingControl', p: { 'w:val': 'compressPunctuation' } },
      { n: 'w:compat', c: [
        { n: 'w:spaceForUL' },
        { n: 'w:balanceSingleByteDoubleByteWidth' },
        { n: 'w:doNotLeaveBackslashAlone' },
        { n: 'w:ulTrailSpace' },
        { n: 'w:doNotExpandShiftReturn' },
        { n: 'w:adjustLineHeightInTable' },
        { n: 'w:useFELayout' },
        { n: 'w:compatSetting', p: { 'w:name': 'compatibilityMode', 'w:uri': 'http://schemas.microsoft.com/office/word', 'w:val': '15' } },
        { n: 'w:compatSetting', p: { 'w:name': 'overrideTableStyleFontSizeAndJustification', 'w:uri': 'http://schemas.microsoft.com/office/word', 'w:val': '1' } },
        { n: 'w:compatSetting', p: { 'w:name': 'enableOpenTypeFeatures', 'w:uri': 'http://schemas.microsoft.com/office/word', 'w:val': '1' } },
        { n: 'w:compatSetting', p: { 'w:name': 'doNotFlipMirrorIndents', 'w:uri': 'http://schemas.microsoft.com/office/word', 'w:val': '1' } },
        { n: 'w:compatSetting', p: { 'w:name': 'differentiateMultirowTableHeaders', 'w:uri': 'http://schemas.microsoft.com/office/word', 'w:val': '1' } },
        { n: 'w:compatSetting', p: { 'w:name': 'useWord2013TrackBottomHyphenation', 'w:uri': 'http://schemas.microsoft.com/office/word', 'w:val': '0' } },
      ]},
      { n: 'w:rsids', c: [
        { n: 'w:rsidRoot', p: { 'w:val': '00ED5EAE' } },
        { n: 'w:rsid', p: { 'w:val': '00251A90' } },
        { n: 'w:rsid', p: { 'w:val': '003E10E1' } },
        { n: 'w:rsid', p: { 'w:val': '0040198B' } },
        { n: 'w:rsid', p: { 'w:val': '006A02E3' } },
        { n: 'w:rsid', p: { 'w:val': '00ED5EAE' } },
        { n: 'w:rsid', p: { 'w:val': '00F777F7' } },
      ]},
      { n: 'm:mathPr', c: [
        { n: 'm:mathFont', p: { 'm:val': 'Cambria Math' } },
        { n: 'm:brkBin', p: { 'm:val': 'before' } },
        { n: 'm:brkBinSub', p: { 'm:val': '--' } },
        { n: 'm:smallFrac', p: { 'm:val': '0' } },
        { n: 'm:dispDef' },
        { n: 'm:lMargin', p: { 'm:val': '0' } },
        { n: 'm:rMargin', p: { 'm:val': '0' } },
        { n: 'm:defJc', p: { 'm:val': 'centerGroup' } },
        { n: 'm:wrapIndent', p: { 'm:val': '1440' } },
        { n: 'm:intLim', p: { 'm:val': 'subSup' } },
        { n: 'm:naryLim', p: { 'm:val': 'undOvr' } },
      ]},
      { n: 'w:themeFontLang', p: { 'w:val': 'en-US', 'w:eastAsia': 'zh-CN' } },
      { n: 'w:clrSchemeMapping', p: { 'w:bg1': 'light1', 'w:t1': 'dark1', 'w:bg2': 'dark2', 'w:accent1': 'accent1', 'w:accent2': 'accent2', 'w:accent3': 'accent3', 'w:accent4': 'accent4', 'w:accent5': 'accent5', 'w:accent6': 'accent6', 'w:hyperlink': 'hyperlink', 'w:followedHyperlink': 'followedHyperlink' } },
      { n: 'w:decimalSymbol', p: { 'm:val': '.' } },
      { n: 'w:listSeparator', p: { 'm:val': ',' } },
      { n: 'w15:chartTrackingRefBased' },
      { n: 'w15:docId', p: { 'w15:val': '{744DEBF8-5BC5-B543-B844-CF5E17982CF2}' } },
    ]
  }
};

export const generateStylesAst = () => {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" 
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" 
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" 
    xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" 
    xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se w16cid">
    <w:docDefaults>
      <w:rPrDefault>
        <w:rPr>
          <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
          <w:kern w:val="2"/>
          <w:sz w:val="21"/>
          <w:szCs w:val="24"/>
          <w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>
        </w:rPr>
      </w:rPrDefault>
      <w:pPrDefault/>
    </w:docDefaults>
    <w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="0" w:defUnhideWhenUsed="0" w:defQFormat="0" w:count="376">
      <w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>
      <w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>
      <w:lsdException w:name="heading 2" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
      <w:lsdException w:name="heading 3" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
      <w:lsdException w:name="heading 4" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
      <w:lsdException w:name="heading 5" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
      <w:lsdException w:name="heading 6" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
      <w:lsdException w:name="heading 7" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
      <w:lsdException w:name="heading 8" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
      <w:lsdException w:name="heading 9" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
      <w:lsdException w:name="index 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="index 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="index 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="index 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="index 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="index 6" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="index 7" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="index 8" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="index 9" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="toc 1" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="toc 2" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="toc 3" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="toc 4" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="toc 5" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="toc 6" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="toc 7" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="toc 8" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="toc 9" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Normal Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="footnote text" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="annotation text" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="header" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="footer" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="index heading" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="caption" w:semiHidden="1" w:uiPriority="35" w:unhideWhenUsed="1" w:qFormat="1"/>
      <w:lsdException w:name="table of figures" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="envelope address" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="envelope return" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="footnote reference" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="annotation reference" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="line number" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="page number" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="endnote reference" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="endnote text" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="table of authorities" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="macro" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="toa heading" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List Bullet" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List Number" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List Bullet 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List Bullet 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List Bullet 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List Bullet 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List Number 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List Number 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List Number 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List Number 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Title" w:uiPriority="10" w:qFormat="1"/>
      <w:lsdException w:name="Closing" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Signature" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Default Paragraph Font" w:semiHidden="1" w:uiPriority="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Body Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Body Text Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List Continue" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List Continue 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List Continue 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List Continue 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="List Continue 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Message Header" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Subtitle" w:uiPriority="11" w:qFormat="1"/>
      <w:lsdException w:name="Salutation" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Date" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Body Text First Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Body Text First Indent 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Note Heading" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Body Text 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Body Text 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Body Text Indent 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Body Text Indent 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Block Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Hyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="FollowedHyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Strong" w:uiPriority="22" w:qFormat="1"/>
      <w:lsdException w:name="Emphasis" w:uiPriority="20" w:qFormat="1"/>
      <w:lsdException w:name="Document Map" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Plain Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="E-mail Signature" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="HTML Top of Form" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="HTML Bottom of Form" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Normal (Web)" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="HTML Acronym" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="HTML Address" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="HTML Cite" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="HTML Code" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="HTML Definition" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="HTML Keyboard" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="HTML Preformatted" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="HTML Sample" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="HTML Typewriter" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="HTML Variable" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="annotation subject" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="No List" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Outline List 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Outline List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Outline List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Simple 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Simple 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Simple 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Classic 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Classic 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Classic 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Classic 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Colorful 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Colorful 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Colorful 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Columns 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Columns 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Columns 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Columns 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Columns 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Grid 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Grid 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Grid 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Grid 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Grid 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Grid 6" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Grid 7" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Grid 8" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table List 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table List 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table List 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table List 6" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table List 7" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table List 8" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table 3D effects 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table 3D effects 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table 3D effects 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Contemporary" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Elegant" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Professional" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Subtle 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Subtle 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Web 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Web 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Balloon Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Table Grid" w:uiPriority="39"/>
      <w:lsdException w:name="Placeholder Text" w:semiHidden="1"/>
      <w:lsdException w:name="No Spacing" w:uiPriority="1" w:qFormat="1"/>
      <w:lsdException w:name="Light Shading" w:uiPriority="60"/>
      <w:lsdException w:name="Light List" w:uiPriority="61"/>
      <w:lsdException w:name="Light Grid" w:uiPriority="62"/>
      <w:lsdException w:name="Medium Shading 1" w:uiPriority="63"/>
      <w:lsdException w:name="Medium Shading 2" w:uiPriority="64"/>
      <w:lsdException w:name="Medium List 1" w:uiPriority="65"/>
      <w:lsdException w:name="Medium List 2" w:uiPriority="66"/>
      <w:lsdException w:name="Medium Grid 1" w:uiPriority="67"/>
      <w:lsdException w:name="Medium Grid 2" w:uiPriority="68"/>
      <w:lsdException w:name="Medium Grid 3" w:uiPriority="69"/>
      <w:lsdException w:name="Dark List" w:uiPriority="70"/>
      <w:lsdException w:name="Colorful Shading" w:uiPriority="71"/>
      <w:lsdException w:name="Colorful List" w:uiPriority="72"/>
      <w:lsdException w:name="Colorful Grid" w:uiPriority="73"/>
      <w:lsdException w:name="Light Shading Accent 1" w:uiPriority="60"/>
      <w:lsdException w:name="Light List Accent 1" w:uiPriority="61"/>
      <w:lsdException w:name="Light Grid Accent 1" w:uiPriority="62"/>
      <w:lsdException w:name="Medium Shading 1 Accent 1" w:uiPriority="63"/>
      <w:lsdException w:name="Medium Shading 2 Accent 1" w:uiPriority="64"/>
      <w:lsdException w:name="Medium List 1 Accent 1" w:uiPriority="65"/>
      <w:lsdException w:name="Revision" w:semiHidden="1"/>
      <w:lsdException w:name="List Paragraph" w:uiPriority="34" w:qFormat="1"/>
      <w:lsdException w:name="Quote" w:uiPriority="29" w:qFormat="1"/>
      <w:lsdException w:name="Intense Quote" w:uiPriority="30" w:qFormat="1"/>
      <w:lsdException w:name="Medium List 2 Accent 1" w:uiPriority="66"/>
      <w:lsdException w:name="Medium Grid 1 Accent 1" w:uiPriority="67"/>
      <w:lsdException w:name="Medium Grid 2 Accent 1" w:uiPriority="68"/>
      <w:lsdException w:name="Medium Grid 3 Accent 1" w:uiPriority="69"/>
      <w:lsdException w:name="Dark List Accent 1" w:uiPriority="70"/>
      <w:lsdException w:name="Colorful Shading Accent 1" w:uiPriority="71"/>
      <w:lsdException w:name="Colorful List Accent 1" w:uiPriority="72"/>
      <w:lsdException w:name="Colorful Grid Accent 1" w:uiPriority="73"/>
      <w:lsdException w:name="Light Shading Accent 2" w:uiPriority="60"/>
      <w:lsdException w:name="Light List Accent 2" w:uiPriority="61"/>
      <w:lsdException w:name="Light Grid Accent 2" w:uiPriority="62"/>
      <w:lsdException w:name="Medium Shading 1 Accent 2" w:uiPriority="63"/>
      <w:lsdException w:name="Medium Shading 2 Accent 2" w:uiPriority="64"/>
      <w:lsdException w:name="Medium List 1 Accent 2" w:uiPriority="65"/>
      <w:lsdException w:name="Medium List 2 Accent 2" w:uiPriority="66"/>
      <w:lsdException w:name="Medium Grid 1 Accent 2" w:uiPriority="67"/>
      <w:lsdException w:name="Medium Grid 2 Accent 2" w:uiPriority="68"/>
      <w:lsdException w:name="Medium Grid 3 Accent 2" w:uiPriority="69"/>
      <w:lsdException w:name="Dark List Accent 2" w:uiPriority="70"/>
      <w:lsdException w:name="Colorful Shading Accent 2" w:uiPriority="71"/>
      <w:lsdException w:name="Colorful List Accent 2" w:uiPriority="72"/>
      <w:lsdException w:name="Colorful Grid Accent 2" w:uiPriority="73"/>
      <w:lsdException w:name="Light Shading Accent 3" w:uiPriority="60"/>
      <w:lsdException w:name="Light List Accent 3" w:uiPriority="61"/>
      <w:lsdException w:name="Light Grid Accent 3" w:uiPriority="62"/>
      <w:lsdException w:name="Medium Shading 1 Accent 3" w:uiPriority="63"/>
      <w:lsdException w:name="Medium Shading 2 Accent 3" w:uiPriority="64"/>
      <w:lsdException w:name="Medium List 1 Accent 3" w:uiPriority="65"/>
      <w:lsdException w:name="Medium List 2 Accent 3" w:uiPriority="66"/>
      <w:lsdException w:name="Medium Grid 1 Accent 3" w:uiPriority="67"/>
      <w:lsdException w:name="Medium Grid 2 Accent 3" w:uiPriority="68"/>
      <w:lsdException w:name="Medium Grid 3 Accent 3" w:uiPriority="69"/>
      <w:lsdException w:name="Dark List Accent 3" w:uiPriority="70"/>
      <w:lsdException w:name="Colorful Shading Accent 3" w:uiPriority="71"/>
      <w:lsdException w:name="Colorful List Accent 3" w:uiPriority="72"/>
      <w:lsdException w:name="Colorful Grid Accent 3" w:uiPriority="73"/>
      <w:lsdException w:name="Light Shading Accent 4" w:uiPriority="60"/>
      <w:lsdException w:name="Light List Accent 4" w:uiPriority="61"/>
      <w:lsdException w:name="Light Grid Accent 4" w:uiPriority="62"/>
      <w:lsdException w:name="Medium Shading 1 Accent 4" w:uiPriority="63"/>
      <w:lsdException w:name="Medium Shading 2 Accent 4" w:uiPriority="64"/>
      <w:lsdException w:name="Medium List 1 Accent 4" w:uiPriority="65"/>
      <w:lsdException w:name="Medium List 2 Accent 4" w:uiPriority="66"/>
      <w:lsdException w:name="Medium Grid 1 Accent 4" w:uiPriority="67"/>
      <w:lsdException w:name="Medium Grid 2 Accent 4" w:uiPriority="68"/>
      <w:lsdException w:name="Medium Grid 3 Accent 4" w:uiPriority="69"/>
      <w:lsdException w:name="Dark List Accent 4" w:uiPriority="70"/>
      <w:lsdException w:name="Colorful Shading Accent 4" w:uiPriority="71"/>
      <w:lsdException w:name="Colorful List Accent 4" w:uiPriority="72"/>
      <w:lsdException w:name="Colorful Grid Accent 4" w:uiPriority="73"/>
      <w:lsdException w:name="Light Shading Accent 5" w:uiPriority="60"/>
      <w:lsdException w:name="Light List Accent 5" w:uiPriority="61"/>
      <w:lsdException w:name="Light Grid Accent 5" w:uiPriority="62"/>
      <w:lsdException w:name="Medium Shading 1 Accent 5" w:uiPriority="63"/>
      <w:lsdException w:name="Medium Shading 2 Accent 5" w:uiPriority="64"/>
      <w:lsdException w:name="Medium List 1 Accent 5" w:uiPriority="65"/>
      <w:lsdException w:name="Medium List 2 Accent 5" w:uiPriority="66"/>
      <w:lsdException w:name="Medium Grid 1 Accent 5" w:uiPriority="67"/>
      <w:lsdException w:name="Medium Grid 2 Accent 5" w:uiPriority="68"/>
      <w:lsdException w:name="Medium Grid 3 Accent 5" w:uiPriority="69"/>
      <w:lsdException w:name="Dark List Accent 5" w:uiPriority="70"/>
      <w:lsdException w:name="Colorful Shading Accent 5" w:uiPriority="71"/>
      <w:lsdException w:name="Colorful List Accent 5" w:uiPriority="72"/>
      <w:lsdException w:name="Colorful Grid Accent 5" w:uiPriority="73"/>
      <w:lsdException w:name="Light Shading Accent 6" w:uiPriority="60"/>
      <w:lsdException w:name="Light List Accent 6" w:uiPriority="61"/>
      <w:lsdException w:name="Light Grid Accent 6" w:uiPriority="62"/>
      <w:lsdException w:name="Medium Shading 1 Accent 6" w:uiPriority="63"/>
      <w:lsdException w:name="Medium Shading 2 Accent 6" w:uiPriority="64"/>
      <w:lsdException w:name="Medium List 1 Accent 6" w:uiPriority="65"/>
      <w:lsdException w:name="Medium List 2 Accent 6" w:uiPriority="66"/>
      <w:lsdException w:name="Medium Grid 1 Accent 6" w:uiPriority="67"/>
      <w:lsdException w:name="Medium Grid 2 Accent 6" w:uiPriority="68"/>
      <w:lsdException w:name="Medium Grid 3 Accent 6" w:uiPriority="69"/>
      <w:lsdException w:name="Dark List Accent 6" w:uiPriority="70"/>
      <w:lsdException w:name="Colorful Shading Accent 6" w:uiPriority="71"/>
      <w:lsdException w:name="Colorful List Accent 6" w:uiPriority="72"/>
      <w:lsdException w:name="Colorful Grid Accent 6" w:uiPriority="73"/>
      <w:lsdException w:name="Subtle Emphasis" w:uiPriority="19" w:qFormat="1"/>
      <w:lsdException w:name="Intense Emphasis" w:uiPriority="21" w:qFormat="1"/>
      <w:lsdException w:name="Subtle Reference" w:uiPriority="31" w:qFormat="1"/>
      <w:lsdException w:name="Intense Reference" w:uiPriority="32" w:qFormat="1"/>
      <w:lsdException w:name="Book Title" w:uiPriority="33" w:qFormat="1"/>
      <w:lsdException w:name="Bibliography" w:semiHidden="1" w:uiPriority="37" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="TOC Heading" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1" w:qFormat="1"/>
      <w:lsdException w:name="Plain Table 1" w:uiPriority="41"/>
      <w:lsdException w:name="Plain Table 2" w:uiPriority="42"/>
      <w:lsdException w:name="Plain Table 3" w:uiPriority="43"/>
      <w:lsdException w:name="Plain Table 4" w:uiPriority="44"/>
      <w:lsdException w:name="Plain Table 5" w:uiPriority="45"/>
      <w:lsdException w:name="Grid Table Light" w:uiPriority="40"/>
      <w:lsdException w:name="Grid Table 1 Light" w:uiPriority="46"/>
      <w:lsdException w:name="Grid Table 2" w:uiPriority="47"/>
      <w:lsdException w:name="Grid Table 3" w:uiPriority="48"/>
      <w:lsdException w:name="Grid Table 4" w:uiPriority="49"/>
      <w:lsdException w:name="Grid Table 5 Dark" w:uiPriority="50"/>
      <w:lsdException w:name="Grid Table 6 Colorful" w:uiPriority="51"/>
      <w:lsdException w:name="Grid Table 7 Colorful" w:uiPriority="52"/>
      <w:lsdException w:name="Grid Table 1 Light Accent 1" w:uiPriority="46"/>
      <w:lsdException w:name="Grid Table 2 Accent 1" w:uiPriority="47"/>
      <w:lsdException w:name="Grid Table 3 Accent 1" w:uiPriority="48"/>
      <w:lsdException w:name="Grid Table 4 Accent 1" w:uiPriority="49"/>
      <w:lsdException w:name="Grid Table 5 Dark Accent 1" w:uiPriority="50"/>
      <w:lsdException w:name="Grid Table 6 Colorful Accent 1" w:uiPriority="51"/>
      <w:lsdException w:name="Grid Table 7 Colorful Accent 1" w:uiPriority="52"/>
      <w:lsdException w:name="Grid Table 1 Light Accent 2" w:uiPriority="46"/>
      <w:lsdException w:name="Grid Table 2 Accent 2" w:uiPriority="47"/>
      <w:lsdException w:name="Grid Table 3 Accent 2" w:uiPriority="48"/>
      <w:lsdException w:name="Grid Table 4 Accent 2" w:uiPriority="49"/>
      <w:lsdException w:name="Grid Table 5 Dark Accent 2" w:uiPriority="50"/>
      <w:lsdException w:name="Grid Table 6 Colorful Accent 2" w:uiPriority="51"/>
      <w:lsdException w:name="Grid Table 7 Colorful Accent 2" w:uiPriority="52"/>
      <w:lsdException w:name="Grid Table 1 Light Accent 3" w:uiPriority="46"/>
      <w:lsdException w:name="Grid Table 2 Accent 3" w:uiPriority="47"/>
      <w:lsdException w:name="Grid Table 3 Accent 3" w:uiPriority="48"/>
      <w:lsdException w:name="Grid Table 4 Accent 3" w:uiPriority="49"/>
      <w:lsdException w:name="Grid Table 5 Dark Accent 3" w:uiPriority="50"/>
      <w:lsdException w:name="Grid Table 6 Colorful Accent 3" w:uiPriority="51"/>
      <w:lsdException w:name="Grid Table 7 Colorful Accent 3" w:uiPriority="52"/>
      <w:lsdException w:name="Grid Table 1 Light Accent 4" w:uiPriority="46"/>
      <w:lsdException w:name="Grid Table 2 Accent 4" w:uiPriority="47"/>
      <w:lsdException w:name="Grid Table 3 Accent 4" w:uiPriority="48"/>
      <w:lsdException w:name="Grid Table 4 Accent 4" w:uiPriority="49"/>
      <w:lsdException w:name="Grid Table 5 Dark Accent 4" w:uiPriority="50"/>
      <w:lsdException w:name="Grid Table 6 Colorful Accent 4" w:uiPriority="51"/>
      <w:lsdException w:name="Grid Table 7 Colorful Accent 4" w:uiPriority="52"/>
      <w:lsdException w:name="Grid Table 1 Light Accent 5" w:uiPriority="46"/>
      <w:lsdException w:name="Grid Table 2 Accent 5" w:uiPriority="47"/>
      <w:lsdException w:name="Grid Table 3 Accent 5" w:uiPriority="48"/>
      <w:lsdException w:name="Grid Table 4 Accent 5" w:uiPriority="49"/>
      <w:lsdException w:name="Grid Table 5 Dark Accent 5" w:uiPriority="50"/>
      <w:lsdException w:name="Grid Table 6 Colorful Accent 5" w:uiPriority="51"/>
      <w:lsdException w:name="Grid Table 7 Colorful Accent 5" w:uiPriority="52"/>
      <w:lsdException w:name="Grid Table 1 Light Accent 6" w:uiPriority="46"/>
      <w:lsdException w:name="Grid Table 2 Accent 6" w:uiPriority="47"/>
      <w:lsdException w:name="Grid Table 3 Accent 6" w:uiPriority="48"/>
      <w:lsdException w:name="Grid Table 4 Accent 6" w:uiPriority="49"/>
      <w:lsdException w:name="Grid Table 5 Dark Accent 6" w:uiPriority="50"/>
      <w:lsdException w:name="Grid Table 6 Colorful Accent 6" w:uiPriority="51"/>
      <w:lsdException w:name="Grid Table 7 Colorful Accent 6" w:uiPriority="52"/>
      <w:lsdException w:name="List Table 1 Light" w:uiPriority="46"/>
      <w:lsdException w:name="List Table 2" w:uiPriority="47"/>
      <w:lsdException w:name="List Table 3" w:uiPriority="48"/>
      <w:lsdException w:name="List Table 4" w:uiPriority="49"/>
      <w:lsdException w:name="List Table 5 Dark" w:uiPriority="50"/>
      <w:lsdException w:name="List Table 6 Colorful" w:uiPriority="51"/>
      <w:lsdException w:name="List Table 7 Colorful" w:uiPriority="52"/>
      <w:lsdException w:name="List Table 1 Light Accent 1" w:uiPriority="46"/>
      <w:lsdException w:name="List Table 2 Accent 1" w:uiPriority="47"/>
      <w:lsdException w:name="List Table 3 Accent 1" w:uiPriority="48"/>
      <w:lsdException w:name="List Table 4 Accent 1" w:uiPriority="49"/>
      <w:lsdException w:name="List Table 5 Dark Accent 1" w:uiPriority="50"/>
      <w:lsdException w:name="List Table 6 Colorful Accent 1" w:uiPriority="51"/>
      <w:lsdException w:name="List Table 7 Colorful Accent 1" w:uiPriority="52"/>
      <w:lsdException w:name="List Table 1 Light Accent 2" w:uiPriority="46"/>
      <w:lsdException w:name="List Table 2 Accent 2" w:uiPriority="47"/>
      <w:lsdException w:name="List Table 3 Accent 2" w:uiPriority="48"/>
      <w:lsdException w:name="List Table 4 Accent 2" w:uiPriority="49"/>
      <w:lsdException w:name="List Table 5 Dark Accent 2" w:uiPriority="50"/>
      <w:lsdException w:name="List Table 6 Colorful Accent 2" w:uiPriority="51"/>
      <w:lsdException w:name="List Table 7 Colorful Accent 2" w:uiPriority="52"/>
      <w:lsdException w:name="List Table 1 Light Accent 3" w:uiPriority="46"/>
      <w:lsdException w:name="List Table 2 Accent 3" w:uiPriority="47"/>
      <w:lsdException w:name="List Table 3 Accent 3" w:uiPriority="48"/>
      <w:lsdException w:name="List Table 4 Accent 3" w:uiPriority="49"/>
      <w:lsdException w:name="List Table 5 Dark Accent 3" w:uiPriority="50"/>
      <w:lsdException w:name="List Table 6 Colorful Accent 3" w:uiPriority="51"/>
      <w:lsdException w:name="List Table 7 Colorful Accent 3" w:uiPriority="52"/>
      <w:lsdException w:name="List Table 1 Light Accent 4" w:uiPriority="46"/>
      <w:lsdException w:name="List Table 2 Accent 4" w:uiPriority="47"/>
      <w:lsdException w:name="List Table 3 Accent 4" w:uiPriority="48"/>
      <w:lsdException w:name="List Table 4 Accent 4" w:uiPriority="49"/>
      <w:lsdException w:name="List Table 5 Dark Accent 4" w:uiPriority="50"/>
      <w:lsdException w:name="List Table 6 Colorful Accent 4" w:uiPriority="51"/>
      <w:lsdException w:name="List Table 7 Colorful Accent 4" w:uiPriority="52"/>
      <w:lsdException w:name="List Table 1 Light Accent 5" w:uiPriority="46"/>
      <w:lsdException w:name="List Table 2 Accent 5" w:uiPriority="47"/>
      <w:lsdException w:name="List Table 3 Accent 5" w:uiPriority="48"/>
      <w:lsdException w:name="List Table 4 Accent 5" w:uiPriority="49"/>
      <w:lsdException w:name="List Table 5 Dark Accent 5" w:uiPriority="50"/>
      <w:lsdException w:name="List Table 6 Colorful Accent 5" w:uiPriority="51"/>
      <w:lsdException w:name="List Table 7 Colorful Accent 5" w:uiPriority="52"/>
      <w:lsdException w:name="List Table 1 Light Accent 6" w:uiPriority="46"/>
      <w:lsdException w:name="List Table 2 Accent 6" w:uiPriority="47"/>
      <w:lsdException w:name="List Table 3 Accent 6" w:uiPriority="48"/>
      <w:lsdException w:name="List Table 4 Accent 6" w:uiPriority="49"/>
      <w:lsdException w:name="List Table 5 Dark Accent 6" w:uiPriority="50"/>
      <w:lsdException w:name="List Table 6 Colorful Accent 6" w:uiPriority="51"/>
      <w:lsdException w:name="List Table 7 Colorful Accent 6" w:uiPriority="52"/>
      <w:lsdException w:name="Mention" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Smart Hyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Hashtag" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Unresolved Mention" w:semiHidden="1" w:unhideWhenUsed="1"/>
      <w:lsdException w:name="Smart Link" w:semiHidden="1" w:unhideWhenUsed="1"/>
    </w:latentStyles>
    <w:style w:type="paragraph" w:default="1" w:styleId="a">
      <w:name w:val="Normal"/>
      <w:qFormat/>
      <w:pPr>
        <w:widowControl w:val="0"/>
        <w:jc w:val="both"/>
      </w:pPr>
    </w:style>
    <w:style w:type="character" w:default="1" w:styleId="a0">
      <w:name w:val="Default Paragraph Font"/>
      <w:uiPriority w:val="1"/>
      <w:semiHidden/>
      <w:unhideWhenUsed/>
    </w:style>
    <w:style w:type="table" w:default="1" w:styleId="a1">
      <w:name w:val="Normal Table"/>
      <w:uiPriority w:val="99"/>
      <w:semiHidden/>
      <w:unhideWhenUsed/>
      <w:tblPr>
        <w:tblInd w:w="0" w:type="dxa"/>
        <w:tblCellMar>
          <w:top w:w="0" w:type="dxa"/>
          <w:left w:w="108" w:type="dxa"/>
          <w:bottom w:w="0" w:type="dxa"/>
          <w:right w:w="108" w:type="dxa"/>
        </w:tblCellMar>
      </w:tblPr>
    </w:style>
    <w:style w:type="numbering" w:default="1" w:styleId="a2">
      <w:name w:val="No List"/>
      <w:uiPriority w:val="99"/>
      <w:semiHidden/>
      <w:unhideWhenUsed/>
    </w:style>
  </w:styles>`;
};