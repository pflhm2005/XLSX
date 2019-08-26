const LOCAL_FILE_HEADER = "PK\x03\x04";
const CENTRAL_FILE_HEADER = "PK\x01\x02";
const CENTRAL_DIRECTORY_END = "PK\x05\x06";
const ZIP64_CENTRAL_DIRECTORY_LOCATOR = "PK\x06\x07";
const ZIP64_CENTRAL_DIRECTORY_END = "PK\x06\x06";
const DATA_DESCRIPTOR = "PK\x07\x08";

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

/**
 * 服务下面面的映射
 */
const getTypeOf = (input) => {
  if(typeof input === 'string') return 'string';
  else if(Array.isArray(input)) return 'array';
  else if(input instanceof Uint8Array) return 'uint8array';
  else if(input instanceof ArrayBuffer) return 'arraybuffer';
}
const strToAr = (str, ar) => {
  for(let i = 0;i < str.length;i++) ar[i] = str.charCodeAt(i) & 0xff;
  return ar;
}
const arToString = (ar) => {
  let chunk = 65536;
  let result = [];
  let len = ar.length;
  let type = getTypeOf(ar);
  let k = 0;

  while(k < len) {
    let nextIndex = k + chunk;
    if(type === 'array') result.push(String.fromCharCode.apply(null, ar.slice(k, Math.min(nextIndex, len))));
    else result.push(String.fromCharCode.apply(null, ar.subarray(k, Math.min(nextIndex, len))));
    k = nextIndex;
  }
  return result.join('');
}
const arToAr = (from, to) => {
  for(let i = 0;i < from.length;i++) to[i] = from[i];
  return to;
}

let transform = {
  string: {
    array(input) { return strToAr(input, new Array(input.length)); },
    arratbuffer(input) { return strToAr(input, new Uint8Array(input.length)); },
    uint8array(input) { return strToAr(input, new Uint8Array(input.length)); }
  },
  array: {
    string: arToString,
    arraybuffer(input) { return (new Uint8Array(input)).buffer; },
    uint8array(input) { return new Uint8Array(input); }
  },
  arraybuffer: {
    string(input) { return arToString(new Uint8Array(input)); },
    array(input) { return arToAr(new Uint8Array(input), new Array(input.byteLength)); },
    uint8array(input) { return new Uint8Array(input); }
  },
  uint8array: {
    string: arToString,
    array(input) { return arToAr(input, new Array(input.length)); },
    arraybuffer(input) { return input.buffer; },
  }
}

const transformTo = (outputType, input) => {
  if(!outputType) return '';
  let inputType = getTypeOf(input);
  if(inputType === outputType) return input;
  return transform[inputType][outputType](input);
}

const string2buf = (str) => {
  let len = str.length;
  let buf_len = 0;
  // 计算长度
  for(let i = 0;i < len;i++) {
    let c = str.charCodeAt(i);
    // ???
    if((c & 0xfc00) === 0xd800 && (i + 1 < len)) {
      let c2 = str.charCodeAt(i + 1);
      if((c2 & 0xfc00) === 0xdc00) {
        c = 0x10000 + ((c - 0xd800) << 10) + (c2 - 0xdc00);
        i++;
      }
    }
    /**
     * 0x80 => 0 ~ 128 => 0 ~ 1 << 7
     * 0x800 => 128 ~ 2048 => 1 << 7 ~ 1 << 11
     * 0x10000 => 2048 ~ 65536 1 << 11 ~ 1 << 16
     */
    buf_len += c < 0x80 ? 1 : c < 0x800 ? 2 : c < 0x10000 ? 3 : 4;
  }

  let buf = new Uint8Array(buf_len);

  for(let i = 0, j = 0;i < buf_len;j++) {
    let c = str.charCodeAt(j);
    // ???
    if((c & 0xfc00) === 0xd800 && (j + 1 < len)) {
      let c2 = str.charCodeAt(j + 1);
      if((c2 & 0xfc00) === 0xdc00) {
        c = 0x10000 + ((c - 0xd800) << 10) + (c2 - 0xdc00);
        j++;
      }
    }
    if(c < 0x80) buf[i++] = c;
    else if(c < 0x800) {
      // 0xc0 = 1 << 6 + 1 << 7
      buf[i++] = 0xc0 | (c >>> 6);
      // 0x3f = 1 << 6 - 1
      buf[i++] = 0x80 | (c & 0x3f);
    } else if(c < 0x10000) {
      buf[i++] = 0xe0 | (c >>> 12);
      buf[i++] = 0x80 | (c >>> 6 & 0x3f);
      buf[i++] = 0x80 | (c & 0x3f);
    } else {
      buf[i++] = 0xf0 | (c >>> 18);
      buf[i++] = 0x80 | (c >>> 12 & 0x3f);
      buf[i++] = 0x80 | (c >>> 6 & 0x3f);
      buf[i++] = 0x80 | (c & 0x3f);
    }
  }

  return buf;
}

const string2binary = (str) => {
  let result = "";
  for (let i = 0; i < str.length; i++) result += String.fromCharCode(str.charCodeAt(i) & 0xff);
  return result;
};

const decode = () => {};
const utf8decode = () => {};

class ZipObject {
  constructor(name, data, options) {
    this.name = name;
    this.dir = options.dir;
    this.date = options.date;
    this.comment = options.comment;
    this.unixPermissions = options.unixPermissions;
    this.dosPermissions = options.dosPermissions;

    this._data = data;
    this.options = options;

    this._initialMetadata = {
      dir : options.dir,
      date : options.date
    };
  }
  asBinary(asUTF8) {
    let result = this._data;
    if(this.options.base64) result = decode(result);

    if(asUTF8 && this.options.binary) result = utf8decode(result);
    else result = transformTo('string', result);

    if(!asUTF8 && !this.options.binary) result = transformTo('string', string2buf(result));
    return result;
  }
}

class JSZip {
  constructor(data, options = {}) {
    this.files = {};
    this.comment = null;
    this.root = '';
  }
  file(name, data) {
    let type = getTypeOf(data);
    let o = {
      base64: false,
      binary: false,
      dir: false,
      createFolders: false,
      date: new Date(),
      compression: null,
      compressionOptions: null,
      comment: null,
      unixPermissions: null,
      dosPermissions: null,
    };

    if(o.dir) {
      o.base64 = false;
      o.binary = false;
      data = null;
      type = null;
    } else if(type === 'string') {
      if(o.binary && !o.base64 && o.optimizedBinaryString !== true) {
        data = string2binary(data);
      }
    } else {
      o.base64 = false;
      o.binary = true;
      if(type === 'arraybuffer') data = transformTo('uint8array', data);
    }
    
    this.files[name] = new ZipObject(name, data, o);
    return this;
  }
  /**
   * Transform an integer into a string in hexadecimal.
   * @param {number} dec the number to convert.
   * @param {number} bytes the number of bytes to generate.
   * @returns {string} the result.
   */
  decToHex(dec, bytes) {
    let hex = '';
    for(let i = 0;i < bytes;i++){
      hex += String.fromCharCode(dec & 0xff);
      dec >>>= 8;
    }
    return hex;
  }
  generate(options) {
    options = Object.assign({
      base64: true,
      compression: "STORE",
      compressionOptions : null,
      type: "base64",
      platform: "DOS",
      comment: null,
      mimeType: 'application/zip',
      encodeFileName: string2buf
    }, options);
    let zipData = [];
    let localDirLength = 0;
    let centralDirLength = 0;

    Object.keys(this.files).forEach(name => {
      let file = this.files[name];
      let compressionObject = this.generateCompressedObjectFrom(file);

      let zipPart = this.generateZipParts(file, compressionObject, localDirLength, options.platform, options.encodeFileName);
      localDirLength += zipPart.fileRecord.length + compressionObject.compressedSize;
      centralDirLength += zipPart.dirRecord.length;
      zipData.push(zipPart);
    });

    let dirEnd = [
      CENTRAL_DIRECTORY_END,
      "\x00\x00",
      "\x00\x00",
      this.decToHex(zipData.length, 2),
      this.decToHex(zipData.length, 2),
      this.decToHex(centralDirLength, 4),
      this.decToHex(localDirLength, 4),
      this.decToHex(0, 2),
    ].join('');

    let totalLength = localDirLength + centralDirLength + dirEnd.length;
    let writer = new Writer(options.type === 'string' ? null : totalLength);
    let len = zipData.length;
    for(let i = 0;i < len;i++) {
      writer.append(zipData[i].fileRecord);
      writer.append(zipData[i].compressedObject.compressedContent);
    }
    for(let i = 0;i < len;i++) {
      writer.append(zipData[i].dirRecord);
    }
    writer.append(dirEnd);

    return writer.finalize();
  }
  crc32(input, crc = 0) {
    let isArray = getTypeOf(input) !== 'string';
    // -crc-1
    crc ^= -1;
    let len = input.length;
    let x = 0,y = 0,b = 0;
    for(let i = 0;i < len;i++) {
      b = isArray ? input[i] : input.charCodeAt(i);
      y = (crc ^ b) & 0xff;
      x = table[y];
      crc = (crc >>> 8) ^ x;
    }
    return crc ^ (-1);
  }
  generateCompressedObjectFrom(file) {
    let result = {
      compressedSize: 0,
      uncompressedSize: 0,
      crc32: 0,
      compressionMethod: null,
      compressedContent: null,
    };
    let content = file.asBinary(false);
    if(!content || content.length === 0 || file.dir) {
      result.compressedContent = "";
      result.crc32 = 0;
    }
    result.uncompressedSize = content.length;
    result.crc32 = this.crc32(content);
    result.compressedContent = content;

    result.compressedSize = result.compressedContent.length;
    result.compressionMethod = "\x00\x00";
    return result;
  }
  generateZipParts(file, compressedObject, offset, platform, encodeFileName) {
    let utfEncodedFileName = transformTo('string', string2buf(file.name));
    
    let extFileAttr = 0;
    let dir = file.dir;
    if(dir) extFileAttr |= 0x00010;

    // dos
    let versionMadeBy = 0x0014;

    let date = file.date;
    let dosTime = date.getHours();
    dosTime <<= 6;
    dosTime |= date.getMinutes();
    dosTime <<= 5;
    dosTime = dosTime | date.getSeconds() / 2;

    let dosDate = date.getFullYear() - 1980;
    dosDate <<= 4;
    dosDate |= (date.getMonth() + 1);
    dosDate <<= 5;
    dosDate |= date.getDate();
    
    let useUTF8ForFileName = utfEncodedFileName.length !== file.name.length;
    let unicodePathExtraField = '';
    let extraFields = '';
    if(useUTF8ForFileName) {
      unicodePathExtraField = [
        this.decToHex(1, 1),
        this.decToHex(this.crc32(utfEncodedFileName), 4),
        utfEncodedFileName,
      ].join('');

      extraFields = [
        "\x75\x70",
        this.decToHex(unicodePathExtraField.length, 2),
        unicodePathExtraField
      ].join('');
    }

    let header = [
      "\x0A\x00",
      useUTF8ForFileName ? "\x00\x08" : "\x00\x00",
      compressedObject.compressionMethod,
      this.decToHex(dosTime, 2),
      this.decToHex(dosDate, 2),
      this.decToHex(compressedObject.crc32, 4),
      this.decToHex(compressedObject.compressedSize, 4),
      this.decToHex(compressedObject.uncompressedSize, 4),
      this.decToHex(utfEncodedFileName.length, 2),
      this.decToHex(extraFields.length, 2),
    ].join('');

    let fileRecord = [
      LOCAL_FILE_HEADER,
      header,
      utfEncodedFileName,
      extraFields,
    ].join('');

    let dirRecord = [
      CENTRAL_FILE_HEADER,
      this.decToHex(versionMadeBy, 2),
      header,
      this.decToHex(0, 2),
      "\x00\x00",
      "\x00\x00",
      this.decToHex(extFileAttr, 4),
      this.decToHex(offset, 4),
      utfEncodedFileName,
      extraFields,
    ].join('');

    return { fileRecord, dirRecord, compressedObject };
  }
}

class Writer {
  constructor(len = null) {
    this.isString = len === null;
    if(this.isString) this.data = [];
    else {
      this.index = 0;
      this.data = new Uint8Array(len);
    }
  }
  append(input) {
    if(this.isString) {
      this.data.push(transformTo('string', input));
    } else{
      input = transformTo('uint8array', input);
      this.data.set(input, this.index);
      this.index += input.length;
    }
  }
  finalize() {
    return this.isString ? this.data.join('') : this.data;
  }
}

export default JSZip;