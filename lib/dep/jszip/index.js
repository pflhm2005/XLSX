function array2string(ar) {
  const len = ar.length;
  // if (len < 65536) {
  return String.fromCharCode.apply(null, ar);
}

function string2buf(str) {
  let buf, c, c2, m_pos, i, str_len = str.length, buf_len = 0;

  // count binary size
  for (m_pos = 0; m_pos < str_len; m_pos++) {
    c = str.charCodeAt(m_pos);
    if ((c & 0xfc00) === 0xd800 && (m_pos + 1 < str_len)) {
      c2 = str.charCodeAt(m_pos + 1);
      if ((c2 & 0xfc00) === 0xdc00) {
        c = 0x10000 + ((c - 0xd800) << 10) + (c2 - 0xdc00);
        m_pos++;
      }
    }
    buf_len += c < 0x80 ? 1 : c < 0x800 ? 2 : c < 0x10000 ? 3 : 4;
  }

  // allocate buffer
  // if (support.uint8array) {
  //   buf = new Uint8Array(buf_len);
  // } else {
  buf = new Array(buf_len);
  // }

  // convert
  for (i = 0, m_pos = 0; i < buf_len; m_pos++) {
    c = str.charCodeAt(m_pos);
    if ((c & 0xfc00) === 0xd800 && (m_pos + 1 < str_len)) {
      c2 = str.charCodeAt(m_pos + 1);
      if ((c2 & 0xfc00) === 0xdc00) {
        c = 0x10000 + ((c - 0xd800) << 10) + (c2 - 0xdc00);
        m_pos++;
      }
    }
    if (c < 0x80) {
      /* one byte */
      buf[i++] = c;
    } else if (c < 0x800) {
      /* two bytes */
      buf[i++] = 0xC0 | (c >>> 6);
      buf[i++] = 0x80 | (c & 0x3f);
    } else if (c < 0x10000) {
      /* three bytes */
      buf[i++] = 0xE0 | (c >>> 12);
      buf[i++] = 0x80 | (c >>> 6 & 0x3f);
      buf[i++] = 0x80 | (c & 0x3f);
    } else {
      /* four bytes */
      buf[i++] = 0xf0 | (c >>> 18);
      buf[i++] = 0x80 | (c >>> 12 & 0x3f);
      buf[i++] = 0x80 | (c >>> 6 & 0x3f);
      buf[i++] = 0x80 | (c & 0x3f);
    }
  }

  return buf;
};

function string2binary(str) {
  const result = null;
  if (typeof Uint8Array !== "undefined") result = new Uint8Array(str.length);
  else result = new Array(str.length);
  for (let i = 0; i < str.length; i++) {
    result[i] = str.charCodeAt(i) & 0xff;
  }
  return result;
}

function s2ab(s) {
  let buf = new ArrayBuffer(s.length);
  let view = new Uint8Array(buf);
  for (let i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}

const table = [
  0x00000000, 0x77073096, 0xEE0E612C, 0x990951BA,
  0x076DC419, 0x706AF48F, 0xE963A535, 0x9E6495A3,
  0x0EDB8832, 0x79DCB8A4, 0xE0D5E91E, 0x97D2D988,
  0x09B64C2B, 0x7EB17CBD, 0xE7B82D07, 0x90BF1D91,
  0x1DB71064, 0x6AB020F2, 0xF3B97148, 0x84BE41DE,
  0x1ADAD47D, 0x6DDDE4EB, 0xF4D4B551, 0x83D385C7,
  0x136C9856, 0x646BA8C0, 0xFD62F97A, 0x8A65C9EC,
  0x14015C4F, 0x63066CD9, 0xFA0F3D63, 0x8D080DF5,
  0x3B6E20C8, 0x4C69105E, 0xD56041E4, 0xA2677172,
  0x3C03E4D1, 0x4B04D447, 0xD20D85FD, 0xA50AB56B,
  0x35B5A8FA, 0x42B2986C, 0xDBBBC9D6, 0xACBCF940,
  0x32D86CE3, 0x45DF5C75, 0xDCD60DCF, 0xABD13D59,
  0x26D930AC, 0x51DE003A, 0xC8D75180, 0xBFD06116,
  0x21B4F4B5, 0x56B3C423, 0xCFBA9599, 0xB8BDA50F,
  0x2802B89E, 0x5F058808, 0xC60CD9B2, 0xB10BE924,
  0x2F6F7C87, 0x58684C11, 0xC1611DAB, 0xB6662D3D,
  0x76DC4190, 0x01DB7106, 0x98D220BC, 0xEFD5102A,
  0x71B18589, 0x06B6B51F, 0x9FBFE4A5, 0xE8B8D433,
  0x7807C9A2, 0x0F00F934, 0x9609A88E, 0xE10E9818,
  0x7F6A0DBB, 0x086D3D2D, 0x91646C97, 0xE6635C01,
  0x6B6B51F4, 0x1C6C6162, 0x856530D8, 0xF262004E,
  0x6C0695ED, 0x1B01A57B, 0x8208F4C1, 0xF50FC457,
  0x65B0D9C6, 0x12B7E950, 0x8BBEB8EA, 0xFCB9887C,
  0x62DD1DDF, 0x15DA2D49, 0x8CD37CF3, 0xFBD44C65,
  0x4DB26158, 0x3AB551CE, 0xA3BC0074, 0xD4BB30E2,
  0x4ADFA541, 0x3DD895D7, 0xA4D1C46D, 0xD3D6F4FB,
  0x4369E96A, 0x346ED9FC, 0xAD678846, 0xDA60B8D0,
  0x44042D73, 0x33031DE5, 0xAA0A4C5F, 0xDD0D7CC9,
  0x5005713C, 0x270241AA, 0xBE0B1010, 0xC90C2086,
  0x5768B525, 0x206F85B3, 0xB966D409, 0xCE61E49F,
  0x5EDEF90E, 0x29D9C998, 0xB0D09822, 0xC7D7A8B4,
  0x59B33D17, 0x2EB40D81, 0xB7BD5C3B, 0xC0BA6CAD,
  0xEDB88320, 0x9ABFB3B6, 0x03B6E20C, 0x74B1D29A,
  0xEAD54739, 0x9DD277AF, 0x04DB2615, 0x73DC1683,
  0xE3630B12, 0x94643B84, 0x0D6D6A3E, 0x7A6A5AA8,
  0xE40ECF0B, 0x9309FF9D, 0x0A00AE27, 0x7D079EB1,
  0xF00F9344, 0x8708A3D2, 0x1E01F268, 0x6906C2FE,
  0xF762575D, 0x806567CB, 0x196C3671, 0x6E6B06E7,
  0xFED41B76, 0x89D32BE0, 0x10DA7A5A, 0x67DD4ACC,
  0xF9B9DF6F, 0x8EBEEFF9, 0x17B7BE43, 0x60B08ED5,
  0xD6D6A3E8, 0xA1D1937E, 0x38D8C2C4, 0x4FDFF252,
  0xD1BB67F1, 0xA6BC5767, 0x3FB506DD, 0x48B2364B,
  0xD80D2BDA, 0xAF0A1B4C, 0x36034AF6, 0x41047A60,
  0xDF60EFC3, 0xA867DF55, 0x316E8EEF, 0x4669BE79,
  0xCB61B38C, 0xBC66831A, 0x256FD2A0, 0x5268E236,
  0xCC0C7795, 0xBB0B4703, 0x220216B9, 0x5505262F,
  0xC5BA3BBE, 0xB2BD0B28, 0x2BB45A92, 0x5CB36A04,
  0xC2D7FFA7, 0xB5D0CF31, 0x2CD99E8B, 0x5BDEAE1D,
  0x9B64C2B0, 0xEC63F226, 0x756AA39C, 0x026D930A,
  0x9C0906A9, 0xEB0E363F, 0x72076785, 0x05005713,
  0x95BF4A82, 0xE2B87A14, 0x7BB12BAE, 0x0CB61B38,
  0x92D28E9B, 0xE5D5BE0D, 0x7CDCEFB7, 0x0BDBDF21,
  0x86D3D2D4, 0xF1D4E242, 0x68DDB3F8, 0x1FDA836E,
  0x81BE16CD, 0xF6B9265B, 0x6FB077E1, 0x18B74777,
  0x88085AE6, 0xFF0F6A70, 0x66063BCA, 0x11010B5C,
  0x8F659EFF, 0xF862AE69, 0x616BFFD3, 0x166CCF45,
  0xA00AE278, 0xD70DD2EE, 0x4E048354, 0x3903B3C2,
  0xA7672661, 0xD06016F7, 0x4969474D, 0x3E6E77DB,
  0xAED16A4A, 0xD9D65ADC, 0x40DF0B66, 0x37D83BF0,
  0xA9BCAE53, 0xDEBB9EC5, 0x47B2CF7F, 0x30B5FFE9,
  0xBDBDF21C, 0xCABAC28A, 0x53B39330, 0x24B4A3A6,
  0xBAD03605, 0xCDD70693, 0x54DE5729, 0x23D967BF,
  0xB3667A2E, 0xC4614AB8, 0x5D681B02, 0x2A6F2B94,
  0xB40BBE37, 0xC30C8EA1, 0x5A05DF1B, 0x2D02EF8D
];

const crc32 = (input, crc = 0) => {
  // -crc-1
  crc ^= -1;
  let len = input.length;
  let x = 0, y = 0, b = 0;
  for (let i = 0; i < len; i++) {
    b = input.charCodeAt(i);
    y = (crc ^ b) & 0xff;
    x = table[y];
    crc = (crc >>> 8) ^ x;
  }
  return crc ^ (-1);
}

function decToHex(dec, bytes) {
  let hex = '';
  for (let i = 0; i < bytes; i++) {
    hex += String.fromCharCode(dec & 0xff);
    dec >>>= 8;
  }
  return hex;
}

class JSZip {
  constructor() {
    this.root = '';
    this.files = Object.create(null);
    this.comment = null;

    this.fileData = null;
    this.fileIndex = 0;
    this.contentBuffer = [];
    this.streamInfo = {};
  }
  file(name, data, opt = {}) {
    // 默认fileOpt
    opt = Object.assign({}, {
      base64: false,
      binary: false,
      dir: false,
      createFolders: true,
      date: new Date(),
      compression: null,
      compressionOptions: null,
      comment: null,
      unixPermissions: null,
      dosPermissions: null,
    }, opt);

    // 文件夹属性
    if (opt.dir) {
      opt.base64 = false;
      opt.binary = true;
      data = '';
      opt.compression = 'STORE';
      dataType = 'string';
    }

    this.files[name] = {
      data, ...opt
    };
    return this;
  }
  generate() {
    const opt = {
      streamFiles: false,
      compression: 'STORE',
      compressionOptions: null,
      type: 'string',
      platform: 'DOS',
      comment: null,
      mimeType: 'application/zip',
    };

    let contentArray = [];
    let dirArray = [];
    let centralDirLength = 0;
    let localDirLength = 0;
    let entriesCount = 0;

    /**
     * .zip文件一览图
     * [local file header1]
     * [encryption header1](加密相关 可省略)
     * [file data 1]
     * [data descriptor 1](压缩内容描述信息 仅当文件被压缩时生效)
     * ...
     * [local file headern]
     * [encryption headern]
     * [file data n]
     * [data descriptor n]
     * 
     * [central directory header 1]
     * ...
     * [central directory header n]
     * 
     * [dirEnd]
     */
    Object.keys(this.files).forEach(filename => {
      let file = this.files[filename];

      entriesCount++;
      let dir = file.dir;
      let date = file.date;
      let data = file.data;

      this.streamInfo = {
        compression: {
          magic: "\x00\x00",
        },
        file: {
          name: filename,
          dir, date,
          comment: file.comment || '',
          unixPermissions: file.unixPermissions,
          dosPermissions: file.dosPermissions
        },
        crc32: crc32(data),
        compressedSize: data.length,
        uncompressedSize: data.length
      };

      const record = generateZipParts(this.streamInfo, localDirLength);
      // local header + file data
      contentArray.push(record.fileRecord, data);
      localDirLength += (record.fileRecord.length + data.length);
      centralDirLength += record.dirRecord.length;
      // central directory header
      dirArray.push(record.dirRecord);
    });

    // dirEnd
    const dirEnd = generateCentralDirectoryEnd(entriesCount, centralDirLength, localDirLength);

    contentArray.push(...dirArray, dirEnd);
    return contentArray.join('');
  }
}

const signature = {
  LOCAL_FILE_HEADER: "PK\x03\x04",
  CENTRAL_FILE_HEADER: "PK\x01\x02",
  CENTRAL_DIRECTORY_END: "PK\x05\x06",
  ZIP64_CENTRAL_DIRECTORY_LOCATOR: "PK\x06\x07",
  ZIP64_CENTRAL_DIRECTORY_END: "PK\x06\x06",
  DATA_DESCRIPTOR: "PK\x07\x08"
};

function generateDosExternalFileAttr(dosPermissions, isDir) {
  return (dosPermissions || 0) & 0x3f;
}

/**
 * 生成zip文件/文件夹头部 详细文档见https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT
 * @param {Object} streamInfo 压缩相关参数
 * @param {Number} offset 当前压缩文件内容偏移值
 */
function generateZipParts(streamInfo, offset) {
  const file = streamInfo.file;
  const encodedFileName = array2string(string2buf(file.name))
  const utfEncodedFileName = array2string(string2buf(file.name));
  const comment = file.comment;
  const encodedComment = '';
  const utfEncodedComment = '';
  const useUTF8ForFileName = utfEncodedFileName.length !== file.name.length;
  const useUTF8ForComment = utfEncodedComment.length !== comment.length;

  const dir = file.dir;
  const date = file.date;

  let unicodePathExtraField = '';
  let extraFields = '';

  const compression = streamInfo.compression;

  const dataInfo = {
    crc32: streamInfo.crc32,
    compressedSize: streamInfo.compressedSize,
    uncompressedSize: streamInfo.uncompressedSize
  };

  /**
   * general purpose bit flag
   */
  let bitflag = 0;
  // 表示filename,comment使用uft8来encode
  if (useUTF8ForFileName || useUTF8ForComment) {
    bitflag |= 0x0800;
  }

  /**
   * version made by => 2byte
   * 高位 => 操作系统 00是dos 03是UNIX
   * 低位 => 编码文件的软件版本 0x0014 => 14 => 1.4版本
   * 
   * todo extFileAttr看不懂
   */
  let extFileAttr = 0;
  let versionMadeBy = 0;
  if (dir) {
    extFileAttr |= 0x0010;
  }

  versionMadeBy = 0x0014;
  extFileAttr |= generateDosExternalFileAttr(file.dosPermissions, dir);

  /**
   * date and time fields 见http://www.delorie.com/djgpp/doc/rbinter/it/52/13.html
   * last mod file time => 2bytes
   * last mod file date => 2bytes
   */
  /**
   * last mod file time
   * 简述 => 16位中 15-11存hours 10-5存minutes 4-0存second/2
   * 假设时间为 12时34分56秒 换算后为 01100(12点) 100010(34分) 11100(28秒)
   */
  let dosTime = date.getUTCHours(); // 返回对应时区的小时
  dosTime <<= 6;
  dosTime |= date.getUTCMinutes();
  dosTime <<= 5;
  dosTime |= date.getUTCSeconds() / 2;

  /**
   * last mod file date
   * 简述 => 15-9存year-1980 8-5存month 4-0存day
   * 假设时间为2022年6月29日 换算后为 0100000(32) 0110(6) 11101(29)
   */
  let dosDate = date.getUTCFullYear() - 1980;
  dosDate <<= 4;
  dosDate |= (date.getUTCMonth() + 1);
  dosDate <<= 5;
  dosDate |= date.getUTCDate();

  /**
   * unicode path extra field
   * extraFields => 0x7075 Tsize Version(1) NameCRC32(4) UnicodeName
   */
  if (useUTF8ForFileName) {
    unicodePathExtraField =
      // Version目前是1
      decToHex(1, 1) +
      // NameCRC32
      decToHex(crc32(encodedFileName), 4) +
      // UnicodeName
      utfEncodedFileName;

    extraFields +=
      "\x75\x70" +
      // Tsize
      decToHex(unicodePathExtraField.length, 2) +
      unicodePathExtraField;
  }

  /**
   * Unicode Comment Extra Field
   * 同上 不做处理
   */
  // if (useUTF8ForComment) {}

  /**
   * 文件头部
   * 拼接Local file header 构成如下
   * local file header signature     4 bytes  (0x04034b50)
   * version needed to extract       2 bytes
   * general purpose bit flag        2 bytes
   * compression method              2 bytes
   * last mod file time              2 bytes
   * last mod file date              2 bytes
   * crc-32                          4 bytes
   * compressed size                 4 bytes
   * uncompressed size               4 bytes
   * file name length                2 bytes
   * extra field length              2 bytes
   * file name (variable size)
   * extra field (variable size)
   */
  let header = '';

  // version needed to extract
  // todo 没找到这个的具体定义地点
  header += '\x0A\x00';
  // general purpose bit flag
  header += decToHex(bitflag, 2);
  // compression method
  // 默认是STORE 即不压缩 值为0x0000
  header += compression.magic;
  // last mod file time
  header += decToHex(dosTime, 2);
  // last mod file date
  header += decToHex(dosDate, 2);
  // crc-32
  header += decToHex(dataInfo.crc32, 4);
  // compressed size
  header += decToHex(dataInfo.compressedSize, 4);
  // uncompressed size
  header += decToHex(dataInfo.uncompressedSize, 4);
  // file name length
  header += decToHex(encodedFileName.length, 2);
  // extra field length
  header += decToHex(extraFields.length, 2);

  // 对应Local file header
  const fileRecord = signature.LOCAL_FILE_HEADER + header + encodedFileName + extraFields;

  /**
   * 文件夹头部 
   * central file header signature   4 bytes  (0x02014b50)
   * version made by                 2 bytes
   * version needed to extract       2 bytes
   * general purpose bit flag        2 bytes
   * compression method              2 bytes
   * last mod file time              2 bytes
   * last mod file date              2 bytes
   * crc-32                          4 bytes
   * compressed size                 4 bytes
   * uncompressed size               4 bytes
   * file name length                2 bytes
   * extra field length              2 bytes
   * file comment length             2 bytes
   * disk number start               2 bytes
   * internal file attributes        2 bytes
   * external file attributes        4 bytes
   * relative offset of local header 4 bytes
   * file name (variable size)
   * extra field (variable size)
   * file comment (variable size)
   */
  const dirRecord = signature.CENTRAL_FILE_HEADER +
    decToHex(versionMadeBy, 2) +
    header +
    // file comment length
    decToHex(encodedComment.length, 2) +
    // disk number start
    '\x00\x00' +
    // internal file attributes
    '\x00\x00' +
    // external file attributes
    decToHex(extFileAttr, 4) +
    // relative offset of local header
    decToHex(offset, 4) +
    // file name
    encodedFileName +
    // extra field
    extraFields +
    // file comment
    encodedComment;

  return { fileRecord, dirRecord };
}

function generateCentralDirectoryEnd(entriesCount, centralDirLength, localDirLength, comment) {
  let dirEnd = '';
  if (!comment) comment = '';
  dirEnd = signature.CENTRAL_DIRECTORY_END +
    // number of this disk 磁盘的编号 默认为0x00 zip64位是0xffff
    '\x00\x00' +
    // number of the disk with the start of the central directory 磁盘的编号 默认为0x00 zip64位是0xffff
    '\x00\x00' +
    // total number of entries in the central directory on this disk
    decToHex(entriesCount, 2) +
    // total number of entries in the central directory
    decToHex(entriesCount, 2) +
    // size of the central directory
    decToHex(centralDirLength, 4) +
    // offset of start of central directory with respect to the starting disk number
    decToHex(localDirLength, 4) +
    // .ZIP file comment length
    decToHex(comment.length, 2) +
    // .ZIP file comment
    comment;

  return dirEnd;
}