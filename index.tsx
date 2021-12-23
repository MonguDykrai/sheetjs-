import React, { Component } from 'react';
import { render } from 'react-dom';
import exportToExcel from './export-to-excel';

function App() {
  return (
    <button
      onClick={() => {
        exportToExcel({
          columnWidth: [],
          filename: '工作表' + String(Date.now()),
          header: ['姓名'],
          rows: [{ name: '李雷' }],
        });
      }}
    >
      导出Excel
    </button>
  );
}

render(<App />, document.getElementById('root'));
