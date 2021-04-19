
const readXlsxFile = require('read-excel-file/node');
module.exports = {
  readXlsx: async (xlsxPath) => {
    return await readXlsxFile(xlsxPath, {sheet: 'パワポ入力用'});
  }
}