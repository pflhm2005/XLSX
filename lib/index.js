import XLSX from './xlsx';
import DOCX from './docx';
if (window) window.Office = { XLSX, DOCX };
export {
  XLSX,
  DOCX
}