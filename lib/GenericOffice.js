/* eslint-disable */
const XML_ROOT_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

export default class GenericOffice {
  constructor(type) {
    this.type = type;
    this.a = document.createElement('a');
    this.canvas = document.createElement('canvas');
    this.ctx = this.canvas.getContext("2d");
  }

  s2ab(s) {
    if (typeof s !== 'string') return s;
    let buf = new ArrayBuffer(s.length);
    let view = new Uint8Array(buf);
    for (let i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }

  exportFile(bytes, filename) {
    let blob = new Blob(bytes.map(byte => this.s2ab(byte)), { type: "application/octet-stream" });
    let url = URL.createObjectURL(blob);
    this.a.href = url;
    this.a.download = filename;
    this.a.click();
    setTimeout(() => {
      URL.revokeObjectURL(url);
    }, 10000);
  }

  async getImageData(imageConfig) {
    let imageData = null;
    if (imageConfig.type === 'require') {
      imageData = await this.getImageDataBySrc(imageConfig);
    } else if (imageConfig.type === 'url') {
      imageData = await this.getImageDataByUrl(imageConfig);
    } else if (imageConfig.type === 'upload') {
      imageData = await this.getImageDataByFile(imageConfig);
    }
    return imageData;
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
        })
      };
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
      throw new Error(e);
    });
  }

  writeTag(o) {
    if (!o) return '';
    let properties = '';
    if (o.p) properties = Object.keys(o.p).map(key => ` ${key}="${o.p[key]}"`).join('');
    if (o.t === undefined && !o.c) return `<${o.n}${properties}/>`;
    if (o.c !== undefined) return `<${o.n}${properties}>${(o.c || []).map(v => this.writeTag(v)).join('')}</${o.n}>`;
    return `<${o.n}${properties}>${o.t}</${o.n}>`;
  }

  writeXml(ast) {
    return `${XML_ROOT_HEADER}${this.writeTag(ast)}`;
  }
}
