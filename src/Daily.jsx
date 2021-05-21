/* eslint-disable no-prototype-builtins */
import React from 'react';
import ReactExport from 'react-data-export';
import moment from 'moment';
import XLSX from 'tempa-xlsx';
import JSZip from 'jszip/dist/jszip';
import { saveAs } from 'file-saver';
import data from './data.json';

import ExcelFile from './components/ExcelFile';
import * as _DataUtil from './components/DataUtil';
// const { ExcelFile } = ReactExport;
const { ExcelSheet } = ReactExport.ExcelFile;
const { ExcelColumn } = ReactExport.ExcelFile;
const currency = 'IDR';
const DataByBondType = {};
const DataBySecurities = {};
const compiledData = [
  [{ value: 'Bond Reporting Volume Summary by Bond Type', style: { font: { bold: true } } }],
  [
    {
      value: 'No.',
      style: {
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        fill: { patternType: 'solid', fgColor: { rgb: '800E00' } },
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value: 'Bond Type',
      style: {
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        fill: { patternType: 'solid', fgColor: { rgb: '800E00' } },
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value: 'Today',
      style: {
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        fill: { patternType: 'solid', fgColor: { rgb: '800E00' } },
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value: 'Previous',
      style: {
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        fill: { patternType: 'solid', fgColor: { rgb: '800E00' } },
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value: 'Chg%',
      style: {
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        fill: { patternType: 'solid', fgColor: { rgb: '800E00' } },
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value: 'Daily Average',
      style: {
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        fill: { patternType: 'solid', fgColor: { rgb: '800E00' } },
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value: 'Chg%',
      style: {
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        fill: { patternType: 'solid', fgColor: { rgb: '800E00' } },
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
  ],
];

const dataFiltered = data.filter(
  d => d.buysell === 'S' && d.trade_status === 'M' && d.transaction_currency === currency
);

dataFiltered.forEach(d => {
  if (!DataByBondType[d.securities_type_id]) {
    const outrightTrx = dataFiltered.filter(
      f => f.securities_type_id === d.securities_type_id && f.transaction_type === 'O'
    );
    const nonOutrightTrx = dataFiltered.filter(
      f => f.securities_type_id === d.securities_type_id && f.transaction_type !== 'O'
    );

    DataByBondType[d.securities_type_id] = {};

    DataByBondType[d.securities_type_id].name = d.securities_type;
    DataByBondType[d.securities_type_id].outright = outrightTrx;
    DataByBondType[d.securities_type_id].regular = nonOutrightTrx;
  }
});

dataFiltered.forEach(d => {
  if (!DataBySecurities[d.securities_type_id]) {
    DataBySecurities[d.securities_type_id] = {};
    DataBySecurities[d.securities_type_id].name = d.securities_type;
  }

  if (!DataBySecurities[d.securities_type_id][d.securities_id]) {
    DataBySecurities[d.securities_type_id][d.securities_id] = dataFiltered.filter(
      f => f.securities_type_id === d.securities_type_id && f.securities_id === d.securities_id
    );
  }
});

let num = 1;
let todayVolumeSummary = 0;
let todayFrequencySummary = 0;
let prevVolumeSummary = 0;
let prevFrequencySummary = 0;
let chgSummary = 0;
let avgSummary = 0;
let chg2Summary = 0;

Object.keys(DataByBondType).forEach(d => {
  let todayVolumeTotal = 0;
  let todayFrequencyTotal = 0;
  let todayVolumeOutright = 0;
  let todayFrequencyOutright = 0;
  let todayVolumeNonOutright = 0;
  let todayFrequencyNonOutright = 0;
  let prevVolumeTotal = 0;
  let prevFrequencyTotal = 0;
  let prevVolumeOutright = 0;
  let prevFrequencyOutright = 0;
  let prevVolumeNonOutright = 0;
  let prevFrequencyNonOutright = 0;
  let chg = 0;
  let chgOutright = 0;
  let chgNonOutright = 0;
  let avg = 0;
  let avgOutright = 0;
  let avgNonOutright = 0;
  let chg2 = 0;
  let chg2Outright = 0;
  let chg2NonOutright = 0;

  DataByBondType[d].outright
    .filter(d => moment(d.trade_date).diff(moment(), 'day') === 0)
    ?.forEach(a => {
      todayVolumeOutright += parseFloat(a.volume);
      todayFrequencyOutright += 1;
    });

  DataByBondType[d].outright
    .filter(d => moment(d.trade_date).diff(moment(), 'day') === -1)
    ?.forEach(a => {
      prevVolumeOutright += parseFloat(a.volume);
      prevFrequencyOutright += 1;
    });

  DataByBondType[d].regular
    .filter(d => moment(d.trade_date).diff(moment(), 'day') === 0)
    ?.forEach(a => {
      todayVolumeNonOutright += parseFloat(a.volume);
      todayFrequencyNonOutright += 1;
    });

  DataByBondType[d].regular
    .filter(d => moment(d.trade_date).diff(moment(), 'day') === -1)
    ?.forEach(a => {
      prevVolumeNonOutright += parseFloat(a.volume);
      prevFrequencyNonOutright += parseFloat(a.volume);
    });

  todayVolumeTotal = todayVolumeOutright + todayVolumeNonOutright;
  prevVolumeTotal = prevVolumeOutright + prevVolumeNonOutright;
  todayFrequencyTotal = todayFrequencyOutright + todayFrequencyNonOutright;
  prevFrequencyTotal = prevFrequencyOutright + prevFrequencyNonOutright;
  chg = (todayVolumeTotal - prevVolumeTotal) / prevVolumeTotal;
  avg = todayVolumeTotal / prevFrequencyTotal;
  chg2 =
    (((todayVolumeTotal / todayFrequencyTotal - prevVolumeTotal / prevFrequencyTotal) /
      (prevVolumeTotal / prevFrequencyTotal)) *
      100) /
    100;
  chgOutright = (todayVolumeOutright - prevVolumeOutright) / prevVolumeOutright;
  avgOutright = todayVolumeOutright / prevFrequencyOutright;
  chg2Outright =
    (((todayVolumeOutright / todayFrequencyOutright - prevVolumeOutright / prevFrequencyOutright) /
      (prevVolumeOutright / prevFrequencyOutright)) *
      100) /
    100;

  chgNonOutright = (todayVolumeNonOutright - prevVolumeNonOutright) / prevVolumeNonOutright;
  avgNonOutright = todayVolumeNonOutright / prevFrequencyNonOutright;
  chg2NonOutright =
    (((todayVolumeNonOutright / todayFrequencyNonOutright - prevVolumeNonOutright / prevFrequencyNonOutright) /
      (prevVolumeNonOutright / prevFrequencyNonOutright)) *
      100) /
    100;

  todayVolumeSummary += todayVolumeTotal;
  prevVolumeSummary += prevVolumeTotal;
  todayFrequencySummary += todayFrequencyTotal;
  prevFrequencySummary += prevFrequencyTotal;

  compiledData.push([
    {
      value: num,
      style: {
        font: { bold: true },
        alignment: {
          horizontal: 'center',
          vertical: 'center',
        },
      },
      merge: { r: 2 },
    },
    {
      value: DataByBondType[d].name,
      style: {
        font: { bold: true },
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value: todayVolumeTotal.toLocaleString('en', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
      style: {
        font: { bold: true },
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value: prevVolumeTotal.toLocaleString('en', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
      style: {
        font: { bold: true },
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value:
        chg && prevVolumeTotal > 0
          ? `${chg.toLocaleString('en', {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            })}%`
          : '-',
      style: {
        font: { bold: true },
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value:
        avg && prevFrequencyTotal > 0
          ? avg.toLocaleString('en', {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            })
          : '-',
      style: {
        font: { bold: true },
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value: chg2
        ? `${chg2.toLocaleString('en', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
          })}%`
        : '-',
      style: {
        font: { bold: true },
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
  ]);

  compiledData.push([
    { value: '' },
    {
      value: '- Outright',
      style: {
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value: todayVolumeOutright.toLocaleString('en', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
      style: {
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value: prevVolumeOutright.toLocaleString('en', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
      style: {
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value:
        chgOutright && prevVolumeOutright > 0
          ? `${chgOutright.toLocaleString('en', {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            })}%`
          : '-',
      style: {
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value:
        avgOutright && prevFrequencyOutright > 0
          ? avgOutright.toLocaleString('en', {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            })
          : '-',
      style: {
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value: chg2Outright
        ? `${chg2Outright.toLocaleString('en', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
          })}%`
        : '-',
      style: {
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
  ]);

  compiledData.push([
    { value: '' },
    {
      value: '- Non Outright',
      style: {
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value: todayVolumeNonOutright.toLocaleString('en', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
      style: {
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value: prevVolumeNonOutright.toLocaleString('en', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
      style: {
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value:
        chgNonOutright && prevVolumeNonOutright > 0
          ? `${chgNonOutright.toLocaleString('en', {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            })}%`
          : '-',
      style: {
        font: { bold: true },
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value:
        avgNonOutright && prevFrequencyNonOutright > 0
          ? avgNonOutright.toLocaleString('en', {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            })
          : '-',
      style: {
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
    {
      value: chg2NonOutright
        ? `${chg2NonOutright.toLocaleString('en', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
          })}%`
        : '-',
      style: {
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    },
  ]);

  num += 1;
});

chgSummary = (todayVolumeSummary - prevVolumeSummary) / prevVolumeSummary;
avgSummary = todayVolumeSummary / prevFrequencySummary;
chg2Summary =
  (((todayVolumeSummary / todayFrequencySummary - prevVolumeSummary / prevFrequencySummary) /
    (prevVolumeSummary / prevFrequencySummary)) *
    100) /
  100;
compiledData.push([
  {
    value: 'Total',
    style: {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { patternType: 'solid', fgColor: { rgb: '800E00' } },
      border: {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
      },
    },
    merge: {
      c: 1,
    },
  },
  {
    value: '',
    style: {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { patternType: 'solid', fgColor: { rgb: '800E00' } },
      border: {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
      },
    },
  },
  {
    value: todayVolumeSummary.toLocaleString('en', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
    style: {
      font: { bold: true },
      border: {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
      },
    },
  },
  {
    value: prevVolumeSummary.toLocaleString('en', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
    style: {
      font: { bold: true },
      border: {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
      },
    },
  },
  {
    value: chgSummary
      ? `${chgSummary.toLocaleString('en', {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        })}%`
      : '-',
    style: {
      font: { bold: true },
      border: {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
      },
    },
  },
  {
    value: avgSummary
      ? avgSummary.toLocaleString('en', {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        })
      : '-',
    style: {
      font: { bold: true },
      border: {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
      },
    },
  },
  {
    value: chg2Summary
      ? `${chg2Summary.toLocaleString('en', {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        })}%`
      : '-',
    style: {
      font: { bold: true },
      border: {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
      },
    },
  },
]);

const sheet1 = [
  {
    columns: [],
    data: [
      [
        {
          value: 'Beneficiary of Securities Transaction Report (PLTE)',
          style: { font: { bold: true }, alignment: { horizontal: 'center', vertical: 'center' } },
          merge: { c: 6 },
        },
      ],
      [
        {
          value: 'Daily Report Of Bond Transaction (in Bio)',
          style: { font: { bold: true }, alignment: { horizontal: 'center', vertical: 'center' } },
          merge: { c: 6 },
        },
      ],
      [
        {
          value: 'July 21, 2020',
          style: { font: { bold: true }, alignment: { horizontal: 'center', vertical: 'center' } },
          merge: { c: 6 },
        },
      ],
    ],
  },
  {
    columns: [
      { title: '', width: { wpx: 80 } },
      { title: '', width: { wpx: 180 } },
      { title: '', width: { wpx: 100 } },
      { title: '', width: { wpx: 100 } },
      { title: '', width: { wpx: 100 } },
    ],
    data: compiledData,
  },
];

Object.keys(DataBySecurities).forEach(d => {
  const compiledData2 = [
    [{ value: `Top 10 ${DataBySecurities[d].name} Bonds by Volume`, style: { font: { bold: true } } }],
    [
      {
        value: 'No.',
        style: {
          font: { bold: true, color: { rgb: 'FFFFFF' } },
          fill: { patternType: 'solid', fgColor: { rgb: '800E00' } },
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
          },
        },
      },
      {
        value: 'Securities ID',
        style: {
          font: { bold: true, color: { rgb: 'FFFFFF' } },
          fill: { patternType: 'solid', fgColor: { rgb: '800E00' } },
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
          },
        },
      },
      {
        value: 'High',
        style: {
          font: { bold: true, color: { rgb: 'FFFFFF' } },
          fill: { patternType: 'solid', fgColor: { rgb: '800E00' } },
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
          },
        },
      },
      {
        value: 'Low',
        style: {
          font: { bold: true, color: { rgb: 'FFFFFF' } },
          fill: { patternType: 'solid', fgColor: { rgb: '800E00' } },
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
          },
        },
      },
      {
        value: 'Last',
        style: {
          font: { bold: true, color: { rgb: 'FFFFFF' } },
          fill: { patternType: 'solid', fgColor: { rgb: '800E00' } },
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
          },
        },
      },
      {
        value: 'Volume',
        style: {
          font: { bold: true, color: { rgb: 'FFFFFF' } },
          fill: { patternType: 'solid', fgColor: { rgb: '800E00' } },
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
          },
        },
      },
      {
        value: 'Frequency',
        style: {
          font: { bold: true, color: { rgb: 'FFFFFF' } },
          fill: { patternType: 'solid', fgColor: { rgb: '800E00' } },
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
          },
        },
      },
    ],
  ];

  const topten = [];

  Object.keys(DataBySecurities[d]).forEach(d2 => {
    if (Array.isArray(DataBySecurities[d][d2])) {
      const trx = DataBySecurities[d][d2].filter(d => moment(d.trade_date).diff(moment(), 'day') === 0);

      trx.sort((a, b) =>
        moment(a.trade_date) > moment(b.trade_date) ? -1 : moment(b.trade_date) > moment(a.trade_date) ? 1 : 0
      );

      topten.push({
        securitiesId: d2,
        high: Math.max(...trx.map(d => parseFloat(d.price))),
        low: Math.min(...trx.map(d => parseFloat(d.price))),
        last: trx[0].price,
        volume: trx.reduce((prev, { volume }) => prev + parseFloat(volume), 0),
        frequency: trx.length,
      });
    }
  });

  topten.sort((a, b) => (a.volume > b.volume ? -1 : b.volume > a.volume ? 1 : 0));

  topten.slice(0, 10).forEach((d, i) => {
    compiledData2.push([
      {
        value: i + 1,
        style: {
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
          },
        },
      },
      {
        value: d.securitiesId,
        style: {
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
          },
        },
      },
      {
        value: d.high,
        style: {
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
          },
        },
      },
      {
        value: d.low,
        style: {
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
          },
        },
      },
      {
        value: d.last,
        style: {
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
          },
        },
      },
      {
        value: d.volume,
        style: {
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
          },
        },
      },
      {
        value: d.frequency,
        style: {
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
          },
        },
      },
    ]);
  });

  sheet1.push({
    ySteps: 1,
    columns: [],
    data: compiledData2,
  });
});

const sheet2 = [
  {
    name: 'Ardy',
  },
  {
    name: 'Maul',
  },
];

class Download extends React.Component {
  lev1 = data => <table border="1">{data.map(d => this.lev2(d))}</table>;

  lev2 = data => (
    <tr>
      {data.map(d => (
        <td>{d.value}</td>
      ))}
    </tr>
  );

  handleProcessData = () => {
    const dataSheet = (0, _DataUtil.excelSheetFromDataSet)(sheet1);
    console.log(dataSheet);
    const zip = new JSZip();
    const dailyWorkbook = {
      SheetNames: ['Daily Management Report'],
      Sheets: {
        'Daily Management Report': dataSheet,
      },
    };
    const wDaily = XLSX.write(dailyWorkbook, { bookType: 'xlsx', bookSST: true, type: 'binary' });
    zip.file('daily.xlsx', wDaily, { binary: true });
    // const monthlyWorkbook = {
    //   SheetNames: ['Monthly Management Report'],
    //   Sheets: {
    //     'Monthly Management Report': dataSheet,
    //   },
    // };
    // const wMonthly = XLSX.write(monthlyWorkbook, { bookType: 'xlsx', bookSST: true, type: 'binary' });
    // zip.file('monthly.xlsx', wMonthly, { binary: true });

    zip.generateAsync({ type: 'blob' }).then(content => {
      saveAs(content, 'reporting.zip');
    });
  };

  render() {
    return (
      <>
        <button type="button" onClick={this.handleProcessData}>
          download
        </button>
        {sheet1.map(d => this.lev1(d.data))}
        <ExcelFile filename="Daily Management Report">
          <ExcelSheet dataSet={sheet1} name="Daily Management Report" />
          <ExcelSheet data={sheet2} name="Test">
            <ExcelColumn label="Name" value="name" />
          </ExcelSheet>
        </ExcelFile>
      </>
    );
  }
}

export default Download;
