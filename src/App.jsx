import React, { useState } from 'react';
import Daily from './Daily';
import Mass from './Mass';
import MassGov from './MassGov';
import MassCor from './MassCor';
import MassSum from './MassSum';
import Top15Gov from './Top15Gov';
import Top15Cor from './Top15Cor';

export default function App() {
  const [page, setPage] = useState('masssum');

  const renderPage = param => {
    switch (param) {
      case 'daily':
        return <Daily />;
      case 'mass':
        return <Mass />;
      case 'massgov':
        return <MassGov />;
      case 'masscor':
        return <MassCor />;
      case 'masssum':
        return <MassSum />;
      case 'top15gov':
        return <Top15Gov />;
      case 'top15cor':
        return <Top15Cor />;
      default:
        return null;
    }
  };

  return (
    <div>
      <button type="button" onClick={() => setPage('daily')}>
        Daily
      </button>
      <button type="button" onClick={() => setPage('mass')}>
        Mass
      </button>
      <button type="button" onClick={() => setPage('massgov')}>
        Mass Government
      </button>
      <button type="button" onClick={() => setPage('masscor')}>
        Mass Corporate
      </button>
      <button type="button" onClick={() => setPage('masssum')}>
        Mass Summary
      </button>
      <button type="button" onClick={() => setPage('top15gov')}>
        Top 15 PLTE Government
      </button>
      <button type="button" onClick={() => setPage('top15cor')}>
        Top 15 PLTE Corporate
      </button>
      {renderPage(page)}
    </div>
  );
}
