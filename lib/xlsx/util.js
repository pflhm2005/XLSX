export const columnToNum = str => str.split('').reduce((r, c) => {
  return r = r * 26 + (c.charCodeAt() - 64);
}, 0);

export const numToColumn = (n) => {
  let s = '';
  for(++n;n;n = Math.floor((n - 1) / 26)) s = String.fromCharCode((n - 1) % 26 + 65) + s;
  return s;
}

/**
 * 0,0 => A1
 * 0,1 => B1
 * 1,0 => A2
 * @param {Number} r 行
 * @param {Number} c 列
 * @returns 返回对应的坐标
 */
export const transferCellPos = (r, c) => {
  return `${numToColumn(c)}${r + 1}`;
}

export const escapeHTML = (str) => {
  return str.replace(/[<>"&]/g, (match) => {
    switch(match){
      case "<": return "&lt;"; 
      case ">": return "&gt;";
      case "&": return "&amp;"; 
      case "\"": return "&quot;"; 
    };
  })
}