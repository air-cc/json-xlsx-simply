const XLSX = require('xlsx')
const promisify = require('bluebird').promisify

const xlsxWriteFileAsync = promisify(XLSX.writeFileAsync)

module.exports = outputXLSX
module.exports.jsonToXLSX = jsonToXLSX

/**
 * 导出 xlsx
 * 
 * @param {String} filePath
 * @param {String} sheetName 
 * @param {Array} rows   格式：[[head-1, head-2], [value-1, value-2], [value-11, value-22]]
 * @returns 
 */
async function outputXLSX(filePath, sheetName, rows) {
  function sheet_from_array_of_arrays(data, opts) {
    var ws = {};
    var range = {
      s: {
        c: 10000000,
        r: 10000000
      },
      e: {
        c: 0,
        r: 0
      }
    };
    for (var R = 0; R != data.length; ++R) {
      for (var C = 0; C != data[R].length; ++C) {
        if (range.s.r > R) range.s.r = R;
        if (range.s.c > C) range.s.c = C;
        if (range.e.r < R) range.e.r = R;
        if (range.e.c < C) range.e.c = C;
        var cell = {
          v: data[R][C]
        };
        if (cell.v == null) continue;
        var cell_ref = XLSX.utils.encode_cell({
          c: C,
          r: R
        });

        if (typeof cell.v === 'number') cell.t = 'n';
        else if (typeof cell.v === 'boolean') cell.t = 'b';
        else if (cell.v instanceof Date) {
          cell.t = 'n';
          cell.z = XLSX.SSF._table[14];
          cell.v = datenum(cell.v);
        } else cell.t = 's';

        ws[cell_ref] = cell;
      }
    }
    if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
    return ws;
  }

  const wb = {
    SheetNames: [],
    Sheets: {}
  }
  wb.SheetNames.push(sheetName)
  wb.Sheets[sheetName] = sheet_from_array_of_arrays(rows)

  return xlsxWriteFileAsync(filePath, wb)
}

/**
 * json to xlsx
 * 
 * @param {String} filePath 
 * @param {Array} data 
 * @param {Object} opts   {head, sheetName, fields}
 * @returns 
 */
function jsonToXLSX(filePath, data, opts = {}) {
  const formatData = (data, head, fields) => {
    return data.reduce((items, item)=> {
      items.push(fields.map((key)=> item[key]))
      return items
    }, [head])
  }
  
  const keys = Object.keys(data[0])
  const {fields = keys, head = keys, sheetName = 'sheet'} = opts
  const rows = formatData(data, head, fields)  
  return outputXLSX(filePath, sheetName, rows)
}