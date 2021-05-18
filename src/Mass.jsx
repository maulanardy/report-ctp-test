/* eslint-disable no-prototype-builtins */
import React from 'react';
import ReactExport from 'react-data-export';
import moment from 'moment';
import data from './data.json';

const currency = 'IDR';
const { ExcelFile } = ReactExport;
const { ExcelSheet } = ReactExport.ExcelFile;
const dataFiltered = data.filter(
  d => d.buysell === 'S' && d.trade_status === 'M' && d.transaction_currency === currency
);
const compiledData = [
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
      value: 'Securities Name',
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
      value: 'Trade Date',
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
      value: 'Settle Date',
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
      value: 'Repo (Term)',
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
      value: 'Price',
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
      value: `Volume (bio) ${currency}`,
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
      value: `Value (bio) ${currency}`,
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
      value: 'Yield',
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
      value: 'Coupon',
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
      value: 'Rating',
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

dataFiltered
  .filter(d => moment(d.trade_date).diff(moment(), 'day') === 0)
  .forEach((d, i) => {
    compiledData.push([
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
        value: d.securities_name,
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
        value: moment(d.trade_date).format('DD-MMM-YYYY'),
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
        value: moment(d.trade_date).format('DD-MMM-YYYY'),
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
        value: d.transaction_type_desc,
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
        value: '',
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
        value: parseFloat(d.price).toLocaleString('en', {
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
        value: parseFloat(d.volume).toLocaleString('en', {
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
      {
        value: parseFloat(d.yield).toLocaleString('en', {
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
        value: parseFloat(d.coupon_rate).toLocaleString('en', {
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
        value: d.rating,
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

const sheet1 = [
  {
    columns: [],
    data: [
      [{ value: 'Laporan Transaksi Efek Bersifat Utang dan Sukuk ke BEI', style: { font: { bold: true, sz: 14 } } }],
      [{ value: moment().format('LLLL'), style: { font: { bold: true, sz: 14 } } }],
      [
        {
          value: `Daftar seluruh transaksi Efek Bersifat Utang dan Sukuk yang dilaporkan melalui BEI ${moment().format(
            'DD MMMM Y'
          )}`,
          style: { font: { bold: true, sz: 14 } },
        },
      ],
      [
        {
          value: `List of Latest Transaction Report ${moment().format('LL')}`,
          style: { font: { bold: true, sz: 11 } },
        },
      ],
    ],
  },
  {
    columns: [
      { title: '', width: { wpx: 80 } },
      { title: '', width: { wpx: 480 } },
      { title: '', width: { wpx: 100 } },
      { title: '', width: { wpx: 100 } },
      { title: '', width: { wpx: 100 } },
    ],
    data: compiledData,
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

  render() {
    return (
      <>
        {sheet1.map(d => this.lev1(d.data))}
        <ExcelFile filename="All Transaction for Mass Media">
          <ExcelSheet dataSet={sheet1} name="All Transaction for Mass Media" />
        </ExcelFile>
      </>
    );
  }
}

export default Download;
