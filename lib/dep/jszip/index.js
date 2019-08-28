import {
  crc32,
  transformTo,
  string2buf,
  decToHex,
} from './util';

const LOCAL_FILE_HEADER = "PK\x03\x04";
const CENTRAL_FILE_HEADER = "PK\x01\x02";
const CENTRAL_DIRECTORY_END = "PK\x05\x06";
const ZIP64_CENTRAL_DIRECTORY_LOCATOR = "PK\x06\x07";
const ZIP64_CENTRAL_DIRECTORY_END = "PK\x06\x06";
const DATA_DESCRIPTOR = "PK\x07\x08";

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

class JSZip {
  constructor() {
    this.files = {};
    this.comment = null;
    this.root = '';
  }
  file(name, data) {
    this.files[name] = data;
    return this;
  }
  generate(opt = {type: 'string'}) {
    let zipData = [];
    let localDirLength = 0;
    let centralDirLength = 0;

    Object.keys(this.files).forEach(name => {
      let file = this.files[name];
      let compressionObject = this.generateCompressedObject(file);

      let zipPart = this.generateZipParts(name, compressionObject, localDirLength);
      localDirLength += zipPart.fileRecord.length + compressionObject.compressedSize;
      centralDirLength += zipPart.dirRecord.length;
      zipData.push(zipPart);
    });

    let dirEnd = [
      CENTRAL_DIRECTORY_END,
      "\x00\x00",
      "\x00\x00",
      decToHex(zipData.length, 2),
      decToHex(zipData.length, 2),
      decToHex(centralDirLength, 4),
      decToHex(localDirLength, 4),
      decToHex(0, 2),
    ].join('');

    let totalLength = localDirLength + centralDirLength + dirEnd.length;
    let writer = new Writer(opt.type === 'string' ? null : totalLength);
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
  generateCompressedObject(file) {
    let result = {
      compressedSize: 0,
      uncompressedSize: 0,
      crc32: 0,
      compressionMethod: null,
      compressedContent: null,
    };
    let content = transformTo('string', string2buf(file));

    result.uncompressedSize = content.length;
    result.crc32 = crc32(content);
    result.compressedContent = content;

    result.compressedSize = result.compressedContent.length;
    result.compressionMethod = "\x00\x00";
    return result;
  }
  generateZipParts(name, compressedObject, offset) {
    let utfEncodedFileName = transformTo('string', string2buf(name));

    // dos
    let versionMadeBy = 0x0014;

    let date = new Date();
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
    
    let useUTF8ForFileName = utfEncodedFileName.length !== name.length;
    let unicodePathExtraField = '';
    let extraFields = '';
    if(useUTF8ForFileName) {
      unicodePathExtraField = [
        decToHex(1, 1),
        decToHex(crc32(utfEncodedFileName), 4),
        utfEncodedFileName,
      ].join('');

      extraFields = [
        "\x75\x70",
        decToHex(unicodePathExtraField.length, 2),
        unicodePathExtraField
      ].join('');
    }

    let header = [
      "\x0A\x00",
      useUTF8ForFileName ? "\x00\x08" : "\x00\x00",
      compressedObject.compressionMethod,
      decToHex(dosTime, 2),
      decToHex(dosDate, 2),
      decToHex(compressedObject.crc32, 4),
      decToHex(compressedObject.compressedSize, 4),
      decToHex(compressedObject.uncompressedSize, 4),
      decToHex(utfEncodedFileName.length, 2),
      decToHex(extraFields.length, 2),
    ].join('');

    let fileRecord = [
      LOCAL_FILE_HEADER,
      header,
      utfEncodedFileName,
      extraFields,
    ].join('');

    let dirRecord = [
      CENTRAL_FILE_HEADER,
      decToHex(versionMadeBy, 2),
      header,
      decToHex(0, 2),
      "\x00\x00",
      "\x00\x00",
      decToHex(0, 4),
      decToHex(offset, 4),
      utfEncodedFileName,
      extraFields,
    ].join('');

    return { fileRecord, dirRecord, compressedObject };
  }
}

export default JSZip;