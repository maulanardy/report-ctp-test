/* eslint-disable no-continue */
/* eslint-disable no-unused-vars */
/* eslint-disable prefer-destructuring */
/* eslint-disable no-bitwise */
/* eslint-disable no-plusplus */
/* eslint-disable eqeqeq */
/* eslint-disable import/no-extraneous-dependencies */
/* eslint-disable no-multi-assign */
Object.defineProperty(exports, '__esModule', {
  value: true,
});
exports.excelSheetFromDataSet = exports.excelSheetFromAoA = exports.dateToNumber = exports.strToArrBuffer = undefined;

const _typeof =
  typeof Symbol === 'function' && typeof Symbol.iterator === 'symbol'
    ? obj => typeof obj
    : obj =>
        obj && typeof Symbol === 'function' && obj.constructor === Symbol && obj !== Symbol.prototype
          ? 'symbol'
          : typeof obj;

const _tempaXlsx = require('tempa-xlsx');

const _tempaXlsx2 = _interopRequireDefault(_tempaXlsx);

function _interopRequireDefault(obj) {
  return obj && obj.__esModule ? obj : { default: obj };
}

const strToArrBuffer = function strToArrBuffer(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);

  for (let i = 0; i != s.length; ++i) {
    view[i] = s.charCodeAt(i) & 0xff;
  }

  return buf;
};

const dateToNumber = function dateToNumber(v, date1904) {
  if (date1904) {
    v += 1462;
  }

  const epoch = Date.parse(v);

  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
};

const excelSheetFromDataSet = function excelSheetFromDataSet(dataSet) {
  /*
    Assuming the structure of dataset
    {
        xSteps?: number; //How many cells to skips from left
        ySteps?: number; //How many rows to skips from last data
        columns: [array | string]
        data: [array_of_array | string|boolean|number | CellObject]
        fill, font, numFmt, alignment, and border
    }
     */
  if (dataSet === undefined || dataSet.length === 0) {
    return {};
  }

  const ws = {};
  const range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
  let rowCount = 0;
  const merges = [];

  dataSet.forEach(dataSetItem => {
    const { columns } = dataSetItem;
    const xSteps = typeof dataSetItem.xSteps === 'number' ? dataSetItem.xSteps : 0;
    const ySteps = typeof dataSetItem.ySteps === 'number' ? dataSetItem.ySteps : 0;
    const { data } = dataSetItem;
    if (dataSet === undefined || dataSet.length === 0) {
      return;
    }

    rowCount += ySteps;

    const columnsWidth = [];
    if (columns.length >= 0) {
      columns.forEach((col, index) => {
        const cellRef = _tempaXlsx2.default.utils.encode_cell({ c: xSteps + index, r: rowCount });
        fixRange(range, 0, 0, rowCount, xSteps, ySteps);
        const colTitle = col;
        if ((typeof col === 'undefined' ? 'undefined' : _typeof(col)) === 'object') {
          // colTitle = col.title; //moved to getHeaderCell
          columnsWidth.push(col.width || { wpx: 80 }); /* wch (chars), wpx (pixels) - e.g. [{wch:6},{wpx:50}] */
        }
        getHeaderCell(colTitle, cellRef, ws);
      });

      rowCount += 1;
    }

    if (columnsWidth.length > 0) {
      ws['!cols'] = columnsWidth;
    }

    for (let R = 0; R != data.length; ++R, rowCount++) {
      for (let C = 0; C != data[R].length; ++C) {
        const cellRef = _tempaXlsx2.default.utils.encode_cell({ c: C + xSteps, r: rowCount });
        fixRange(range, R, C, rowCount, xSteps, ySteps);
        getCell(data[R][C], cellRef, ws, C, rowCount);
      }
    }
  });

  if (range.s.c < 10000000) {
    ws['!ref'] = _tempaXlsx2.default.utils.encode_range(range);
  }

  if (merges.length) {
    ws['!merges'] = merges;
  }

  console.log(ws);

  return ws;
};

function getHeaderCell(v, cellRef, ws) {
  const cell = {};
  const headerCellStyle = v.style ? v.style : { font: { bold: true } }; // if style is then use it
  cell.v = v.title;
  cell.t = 's';
  cell.s = headerCellStyle;
  ws[cellRef] = cell;
}

function getCell(v, cellRef, ws, column, row) {
  // assume v is indeed the value. for other cases (object, date...) it will be overriden.
  const cell = { v };
  if (v === null) {
    return;
  }
  const merges = [];

  if (v.merge) {
    if (!Object.prototype.hasOwnProperty.call(ws, '!merges')) {
      ws['!merges'] = [];
    }
    ws['!merges'].push({
      s: { r: row, c: column },
      e: { r: v.merge.r ? v.merge.r + row : row, c: v.merge.c ? v.merge.c + column : column },
    });
  }

  const isDate = v instanceof Date;
  if (!isDate && (typeof v === 'undefined' ? 'undefined' : _typeof(v)) === 'object') {
    cell.s = v.style;
    cell.v = v.value;
    v = v.value;
  }

  if (typeof v === 'number') {
    cell.t = 'n';
  } else if (typeof v === 'boolean') {
    cell.t = 'b';
  } else if (isDate) {
    cell.t = 'n';
    cell.z = _tempaXlsx2.default.SSF._table[14];
    cell.v = dateToNumber(cell.v);
  } else {
    cell.t = 's';
  }
  ws[cellRef] = cell;
}

function fixRange(range, R, C, rowCount, xSteps, ySteps) {
  if (range.s.r > R + rowCount) {
    range.s.r = R + rowCount;
  }

  if (range.s.c > C + xSteps) {
    range.s.c = C + xSteps;
  }

  if (range.e.r < R + rowCount) {
    range.e.r = R + rowCount;
  }

  if (range.e.c < C + xSteps) {
    range.e.c = C + xSteps;
  }
}

const excelSheetFromAoA = function excelSheetFromAoA(data) {
  const ws = {};
  const range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };

  for (let R = 0; R != data.length; ++R) {
    for (let C = 0; C != data[R].length; ++C) {
      if (range.s.r > R) {
        range.s.r = R;
      }

      if (range.s.c > C) {
        range.s.c = C;
      }

      if (range.e.r < R) {
        range.e.r = R;
      }

      if (range.e.c < C) {
        range.e.c = C;
      }

      const cell = { v: data[R][C] };
      if (cell.v === null) {
        continue;
      }

      const cellRef = _tempaXlsx2.default.utils.encode_cell({ c: C, r: R });
      if (typeof cell.v === 'number') {
        cell.t = 'n';
      } else if (typeof cell.v === 'boolean') {
        cell.t = 'b';
      } else if (cell.v instanceof Date) {
        cell.t = 'n';
        cell.z = _tempaXlsx2.default.SSF._table[14];
        cell.v = dateToNumber(cell.v);
      } else {
        cell.t = 's';
      }

      ws[cellRef] = cell;
    }
  }

  if (range.s.c < 10000000) {
    ws['!ref'] = _tempaXlsx2.default.utils.encode_range(range);
  }

  return ws;
};

exports.strToArrBuffer = strToArrBuffer;
exports.dateToNumber = dateToNumber;
exports.excelSheetFromAoA = excelSheetFromAoA;
exports.excelSheetFromDataSet = excelSheetFromDataSet;
