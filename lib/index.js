import XLSX from './xlsx';
import DOCX from './docx';

if(typeof exports === 'object' && typeof module !== 'undefined') module.exports = { XLSX, DOCX };
else window.Office = { XLSX, DOCX };