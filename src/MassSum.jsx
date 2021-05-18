/* eslint-disable no-prototype-builtins */
import React from 'react';
import ReactExport from 'react-data-export';
import moment from 'moment';
import data from './data.json';

const currency = 'IDR';
const { ExcelFile } = ReactExport;
const { ExcelSheet } = ReactExport.ExcelFile;
const DataBySecurities = {};
const dataFiltered = data.filter(
  d => d.buysell === 'S' && d.trade_status === 'M' && d.transaction_currency === currency
);

dataFiltered.forEach(d => {
  if (!DataBySecurities[`${d.securities_id}.${d.transaction_type}`]) {
    DataBySecurities[`${d.securities_id}.${d.transaction_type}`] = dataFiltered.filter(
      f => f.securities_id === d.securities_id && f.transaction_type === d.transaction_type
    );
  }
});

const sheet1 = [
  {
    columns: [],
    data: [
      [{ value: 'Beneficiary of Securities Transaction Report (PLTE)', style: { font: { bold: true } } }],
      [{ value: 'Daily Report Of Bond Transaction (in Bio)', style: { font: { bold: true } } }],
      [{ value: 'July 21, 2020', style: { font: { bold: true } } }],
    ],
  },
];

const compiledData2 = [
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
      value: 'Transaction Type',
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
      value: 'Maturity Date',
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
      value: 'TTM',
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
      value: 'Freq',
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
      value: `Total Volume (bio) ${currency}`,
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
      value: `Total Value (bio) ${currency}`,
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

Object.keys(DataBySecurities).forEach(d2 => {
  if (Array.isArray(DataBySecurities[d2])) {
    const trx = DataBySecurities[d2].filter(d => moment(d.trade_date).diff(moment(), 'day') === 0);

    trx.sort((a, b) =>
      moment(a.trade_date) > moment(b.trade_date) ? -1 : moment(b.trade_date) > moment(a.trade_date) ? 1 : 0
    );

    topten.push({
      securitiesId: trx[0].securities_id,
      transactionType: trx[0].transaction_type_desc,
      maturityDate: trx[0].maturity_date,
      ttm: trx[0].ttm,
      high: Math.max(...trx.map(d => parseFloat(d.price))),
      low: Math.min(...trx.map(d => parseFloat(d.price))),
      last: trx[0].price,
      frequency: trx.length,
      volume: trx.reduce((prev, { volume }) => prev + parseFloat(volume), 0),
      value: trx.reduce((prev, { value }) => prev + parseFloat(value), 0),
    });
  }
});

topten.sort((a, b) => (a.volume > b.volume ? -1 : b.volume > a.volume ? 1 : 0));

topten.forEach((d, i) => {
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
      value: d.transactionType,
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
      value: moment(d.maturityDate).format('D/MM/Y'),
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
      value: parseFloat(d.ttm).toLocaleString('en', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      }),
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
      value: parseFloat(d.value).toLocaleString('en', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      }),
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

class Download extends React.Component {
  lev1 = data => <table border="1">{data.map(d => this.lev2(d))}</table>;

  lev2 = data => (
    <tr>
      {data.map(d => (
        <td>{d.value}</td>
      ))}
    </tr>
  );

  render() {
    return (
      <>
        {sheet1.map(d => this.lev1(d.data))}
        <ExcelFile filename="Trade Summary for Mass Media">
          <ExcelSheet dataSet={sheet1} name="Daily Management Report" />
        </ExcelFile>
      </>
    );
  }
}

export default Download;
