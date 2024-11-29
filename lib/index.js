import XLSX from './xlsx';
import DOCX from './docx';
import JSZIP from './dep/jszip';

if (window) {
  window.Office = { XLSX, DOCX };
  window.JSZIP = JSZIP;
}

export {
  XLSX,
  DOCX
}